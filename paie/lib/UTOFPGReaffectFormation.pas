{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 26/06/2003
Modifié le ... :   /  /
Description .. : Rectification des tables libres formation
Mots clefs ... :
*****************************************************************
PT1 05/01/2005 JL V_60 Reprise depuis session au lieu de stage
PT2 12/10/2006 JL V_70 FQ 13458 Ajout 0 a gauche pour matricule
PT3 21/04/2008 FL V_80 Partage formation
}
unit UTOFPGReaffectFormation;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, ParamSoc,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,
  ed_tools, HStatus, UTOB,
  HQRy, HTB97, Grids, HSysMenu, PGOutils, PGOutilsFormation,LookUp,PGOutils2;

type

  TOF_PGREAFFECTFORMATION = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure OnUpdate; override;
  private
    Grille: THGrid;
    Tob_Grille: Tob;
    Modifier, ZoneAModifier: Boolean;
    TitreColonne: TStringList;
    TEvent: TStringList;
    GlbSql, Pref: string;
    procedure GrillePostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure RecupClauseChampReaffecter(IndiceChamp, Row: integer; var st: string);
    procedure MiseEnFormeGrille;
    procedure MiseEnFormeColGrille(Champ, Suffixe: string);
    procedure OnChangeListeTable(Sender: TObject);
//    procedure OnChangeListeChamp(Sender: TObject);
    procedure OnClickBtnActiveWhere(Sender: TObject);
    procedure OnClickBtnMajValue(Sender: TObject);
//    function RecupTabletteAssocie(Champ: string): string;
    procedure RecopieTablesSalaries(var NbMajOk: integer);
    procedure MiseAjourSelectif(var NbMajOk: integer);
    function RecupSyntaxe(Champ: string): string;
    procedure RespElipsisClick(Sender : TObject);
    procedure ExitEdit(Sender: TObject);
  end;

const
  // libellés des messages
  TexteMessage: array[1..12] of string = (
    {1}'Veuillez saisir un domaine de réaffectation.',
    {2}'Attention,la réaffectation de certaines données seront perdues!#13#10Voulez-vous réactualiser la grille?',
    {3}'Veuillez saisir une valeur à réaffecter.',
    {4}'La zone de réaffectation est inconnue.',
    {5}'La zone de réaffectation n''est pas intégrée à la grille.',
    {6}'Veuillez selectionner un champ libre à réaffecter.',
    {7}'Aucunes réaffectations n''ont été apportés aux données.',
    {8}'Aucunes réaffectations n''ont été sauvegardés.',
    {9}'Attention!Vous demandez à recopier les champs libres de la table salarié#13#10 sur les champs libres des stagiaires.#13#10 Voulez-vous continuer le traitement?',
    {10}'Attention!Vous demandez à mettre à jour les champs libres de la formation.#13#10 Voulez-vous continuer le traitement?',
    {11}'Veuillez cliquer sur le bouton de validation vert pour la recopie des données.',
    {12}''
    );

implementation
uses EntPaie, PGEditOutils, PGEditOutils2, P5Util, PGEdtEtat, P5Def;

procedure TOF_PGREAFFECTFORMATION.GrilleCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if not ZoneAModifier then exit;
  if Pos('modifié', Grille.cells[Grille.col, 0]) < 1 then
  begin
    Cancel := True;
    if (Acol = Grille.ColCount - 1) then
    begin
      // En fin de ligne on se positionne au début de la ligne suivante
      Grille.col := 0;
      if (ARow <> Grille.RowCount - 1) then
        Grille.Row := Grille.Row + 1;
    end
    else
      Grille.Col := Grille.Col + 1;
  end;
end;

procedure TOF_PGREAFFECTFORMATION.GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if not ZoneAModifier then exit;
  if Pos('modifié', Grille.Cells[ACol, 0]) < 1 then
    GridGriseCell(Grille, Acol, Arow, Canvas);
end;

procedure TOF_PGREAFFECTFORMATION.OnArgument(Arguments: string);
var
  Min, Max: string;
  Combo: THValComboBox;
  Btn: TToolBarButton97;
  Num: Integer;
  edit : THedit;
begin
  inherited;
  SetControlVisible('PNMODIF', False);
  Modifier := False;
  ZoneAModifier := False;
  for Num := 1 to VH_Paie.NBFormationLibre do
  begin
    if Num > 8 then Break;
    VisibiliteChampFormation(IntToStr(Num), GetControl('PFO_FORMATION' + IntToStr(Num)), GetControl('TPFO_FORMATION' + IntToStr(Num)));
    SetControlVisible('PFO_FORMATION' + IntToStr(Num) + '_', GetControlVisible('PFO_FORMATION' + IntToStr(Num)));
    SetControlVisible('TLA' + IntToStr(Num), GetControlVisible('PFO_FORMATION' + IntToStr(Num)));
  end;
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PFO_TRAVAILN' + IntToStr(Num)), GetControl('TPFO_TRAVAILN' + IntToStr(Num)));
    SetControlVisible('PFO_TRAVAILN' + IntToStr(Num) + '_', GetControlVisible('PFO_TRAVAILN' + IntToStr(Num)));
    SetControlVisible('TA' + IntToStr(Num), GetControlVisible('PFO_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PFO_CODESTAT'), GetControl('TPFO_CODESTAT'));
  SetControlVisible('PFO_CODESTAT_', getControlVisible('PFO_CODESTAT'));
  SetControlVisible('TA5', getControlVisible('PFO_CODESTAT'));
  //Valeur par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  SetControltext('PFO_SALARIE', Min);
  SetControltext('PFO_SALARIE_', Max);
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  SetControlText('PFO_ETABLISSEMENT', Min);
  SetControlText('PFO_ETABLISSEMENT_', Max);
  Grille := THGrid(GetControl('GRILLE'));
  if Grille <> nil then
  begin
    Grille.PostDrawCell := GrillePostDrawCell;
    Grille.OnCellEnter := GrilleCellEnter;
  end;
  TitreColonne := TStringList.Create;

  SetControlVisible('CKTRAVAILN1', (VH_PAIE.PGLibelleOrgStat1 <> ''));
  SetControlVisible('CKTRAVAILN2', (VH_PAIE.PGLibelleOrgStat2 <> ''));
  SetControlVisible('CKTRAVAILN3', (VH_PAIE.PGLibelleOrgStat3 <> ''));
  SetControlVisible('CKTRAVAILN4', (VH_PAIE.PGLibelleOrgStat4 <> ''));
  SetControlVisible('CKCODESTAT', (VH_PAIE.PGLibCodeStat <> ''));
  SetControlVisible('CKFORMATION1', (VH_PAIE.FormationLibre1 <> ''));
  SetControlVisible('CKFORMATION2', (VH_PAIE.FormationLibre2 <> ''));
  SetControlVisible('CKFORMATION3', (VH_PAIE.FormationLibre3 <> ''));
  SetControlVisible('CKFORMATION4', (VH_PAIE.FormationLibre4 <> ''));
  SetControlVisible('CKFORMATION5', (VH_PAIE.FormationLibre5 <> ''));
  SetControlVisible('CKFORMATION6', (VH_PAIE.FormationLibre6 <> ''));
  SetControlVisible('CKFORMATION7', (VH_PAIE.FormationLibre7 <> ''));
  SetControlVisible('CKFORMATION8', (VH_PAIE.FormationLibre8 <> ''));
  SetControlText('LISTETABLE', 'RECOPIEFOR');
  //Gestionnnaire d'evenement
  Combo := THValComboBox(GetControl('LISTETABLE'));
  if combo <> nil then combo.OnChange := OnChangeListeTable;
  Btn := TToolBarButton97(GetControl('BTNACTIVEWHERE'));
  if btn <> nil then Btn.Onclick := OnClickBtnActiveWhere;
  Btn := TToolBarButton97(GetControl('BTNMAJVALUE'));
  if btn <> nil then Btn.Onclick := OnClickBtnMajValue;
  Edit := THEdit(GetControl('PFO_RESPONSFOR'));
  If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick;
  Edit := THEdit(GetControl('PFO_RESPONSFOR_'));
  If Edit <> Nil then Edit.OnElipsisClick := RespElipsisClick;
  edit := ThEdit(getcontrol('PFO_SALARIE'));
  if Edit <> nil then edit.OnExit := ExitEdit;
  edit := ThEdit(getcontrol('PFO_SALARIE_'));
  if Edit <> nil then edit.OnExit := ExitEdit;
end;

procedure TOF_PGREAFFECTFORMATION.OnClose;
begin
  inherited;
  if TitreColonne <> nil then TitreColonne.Free;
  if Tob_Grille <> nil then Tob_Grille.free;
end;

procedure TOF_PGREAFFECTFORMATION.RecupClauseChampReaffecter(IndiceChamp, Row: integer; var st: string);
begin
  St := '';
  if IndiceChamp = -1 then exit;
  repeat // Recherche par Colonnes des valeurs modifiées
    if (Grille.CellValues[IndiceChamp, Row]) <> (Grille.CellValues[IndiceChamp + 1, Row]) then
    begin
      if St <> '' then St := St + ',';
      St := St + TitreColonne.Strings[IndiceChamp] + '="' + Grille.CellValues[IndiceChamp + 1, Row] + '" ';
      TEvent.Add(Grille.Titres.Strings[IndiceChamp + 1] + ' ancienne valeur : ' + Grille.CellValues[IndiceChamp, Row] + ' nouvelle valeur : ' + Grille.CellValues[IndiceChamp + 1,
        Row]);
    end;
    IndiceChamp := IndiceChamp + 2;
  until IndiceChamp >= Grille.ColCount - 1;
end;

procedure TOF_PGREAFFECTFORMATION.OnChangeListeTable(Sender: TObject);
var
  i: integer;
  SalVisible, FormVisible: Boolean;
begin
  SalVisible := False;
  FormVisible := False;
  if (GetControlText('LISTETABLE') = 'RECOPIEFORM') or (GetControlText('LISTETABLE') = 'LIBREFORM') then FormVisible := True;
  if (GetControlText('LISTETABLE') = 'RECOPIESAL') or (GetControlText('LISTETABLE') = 'LIBRESAL') then SalVisible := True;
  for i := 1 to VH_Paie.PGNbreStatOrg do
  begin
    SetControlVisible('TPFO_TRAVAILN' + IntToStr(i), SalVisible);
    SetControlVisible('PFO_TRAVAILN' + IntToStr(i), SalVisible);
    SetControlVisible('TA' + IntToStr(i), SalVisible);
    SetControlVisible('PFO_TRAVAILN' + IntToStr(i) + '_', SalVisible);
    SetControlVisible('CKTRAVAILN' + IntToStr(i), SalVisible);
  end;
  SetControlVisible('CKETABLISSEMENT', SalVisible);
  SetControlVisible('CKLIBEMPLOIFOR', SalVisible);
  SetControlVisible('CKRESPONSFOR', SalVisible);
  SetControlVisible('CKCODESTAT', SalVisible);
  SetControlVisible('PFO_CODESTAT', SalVisible);
  SetControlVisible('TPFO_CODESTAT', SalVisible);
  SetControlVisible('PFO_CODESTAT_', SalVisible);
  SetControlVisible('TA5', SalVisible);
  for i := 1 to VH_Paie.NBFormationLibre do
  begin
    SetControlVisible('CKFORMATION' + IntToStr(i), FormVisible);
    SetControlVisible('TPFO_FORMATION' + IntToStr(i), FormVisible);
    SetControlVisible('PFO_FORMATION' + IntToStr(i), FormVisible);
    SetControlVisible('PFO_FORMATION' + IntToStr(i) + '_', FormVisible);
    SetControlVisible('TLA' + IntToStr(i), FormVisible);
  end;
  if (GetControlText('LISTETABLE') = 'RECOPIEFORM') or (GetControlText('LISTETABLE') = 'RECOPIESAL') then SetControlEnabled('BValider', True);
end;

{procedure TOF_PGREAFFECTFORMATION.OnChangeListeChamp(Sender: TObject);
var
  champ: string;
begin
  champ := GetControlText('LISTECHAMP');
  SetControlProperty('OLDVALUE', 'DataType', '');
  SetControlProperty('NEWVALUE', 'DataType', '');
  if champ <> '' then
  begin
    Champ := RecupTabletteAssocie(Champ);
    SetControlProperty('NEWVALUE', 'DataType', Champ);
    SetControlProperty('OLDVALUE', 'DataType', Champ);
    SetControlText('NEWVALUE', '');
    SetControlText('OLDVALUE', '');
  end;
end;             }


procedure TOF_PGREAFFECTFORMATION.OnClickBtnActiveWhere(Sender: TObject);
var
  ListeTable, StSql, ChSql, Critere, AutrePref1, AutrePref2: string;
  Q: Tquery;
  Pages: TPageControl;
  i, Col, Row, IndiceChamp: integer;
  T: Tob;
  HMTrad: THSystemMenu;
begin
  Pref := '';
  ListeTable := GetControlText('LISTETABLE');
  if (ListeTable = '') then
  begin
    PGIBox(TexteMessage[1], Ecran.caption);
    exit;
  end;
  if (ListeTable = 'RECOPIESAL') or (ListeTable = 'RECOPIEFORM') then
  begin
    PGIBox(TexteMessage[11], Ecran.caption);
    exit;
  end;
  if Modifier then
    if PGIAsk(TexteMessage[2], Ecran.caption) = MrNo then
      exit;
  ListeTable := GetControlText('LISTETABLE');
  Pref := 'PFO';
  if Pref = 'PFO' then
  begin
    AutrePref1 := 'PSA';
    AutrePref2 := 'PSS'; // PT1
  end;
  MiseEnFormeGrille;
  Pages := TPageControl(GetControl('Pages'));
  if Pages <> nil then Critere := RecupWhereCritere(Pages)
  else Critere := '';
  if (GlbSql <> '') then GlbSql := ConvertPrefixe(GlbSql, AutrePref1, Pref);
  if Critere <> '' then
  begin
    Critere := ConvertPrefixe(Critere, AutrePref1, Pref);
    if Pref = 'PFO' then Critere := Critere + 'AND PFO_DATEFIN>="' + USDateTime(StrToDate(GetControlText('PFO_DATEDEBUT'))) + '" ' +
      'AND PFO_DATEFIN<="' + USDateTime(StrToDate(GetControlText('PFO_DATEFIN'))) + '"';
  end
  else
    if Pref = 'PFO' then Critere := 'WHERE PFO_DATEFIN>="' + USDateTime(StrToDate(GetControlText('PFO_DATEDEBUT'))) + '" ' +
    'AND PFO_DATEFIN<="' + USDateTime(StrToDate(GetControlText('PFO_DATEFIN'))) + '"';
  if GlbSql <> '' then
    ChSql := Pref + '_ETABLISSEMENT,' + Pref + '_SALARIE,' + 'PFO_CODESTAGE,' + 'PFO_ORDRE,' + GlbSql
  else
    ChSql := Pref + '_ETABLISSEMENT,' + Pref + '_SALARIE,PFO_CODESTAGE,PFO_ORDRE ';
  StSql := 'SELECT DISTINCT ' + ChSql + ' FROM FORMATIONS ' + Critere + ' ORDER BY ' + PRef + '_SALARIE';
  Q := OpenSql(StSql, True);
  if Tob_Grille <> nil then
  begin
    Tob_Grille.free;
    Tob_Grille := nil;
  end;
  Tob_Grille := Tob.create('La liste à maj', nil, -1);
  Tob_Grille.LoadDetailDB('La liste à maj', '', '', Q, False);
  Ferme(Q);

  Row := 0;
  Grille.RowCount := Tob_Grille.Detail.Count + 1;
  for I := 0 to Tob_Grille.Detail.Count - 1 do
  begin
    T := Tob_Grille.Detail[i];
    Row := Row + 1;
    IndiceChamp := 999;
    with T do
      for Col := 0 to Grille.ColCount - 1 do
      begin
        if Pos('modifié', Grille.Cells[Col, 0]) < 1 then
        begin
          IndiceChamp := IndiceChamp + 1;
          if GetNomChamp(IndiceChamp) <> '' then
          begin
            if Pos('affecté', Grille.Cells[Col, 0]) >= 1 then
            begin
              Critere := Grille.ColFormats[Col + 1];
              ReadTokenPipe(Critere, '=');
              Grille.Cells[Col, Row] := RechDom(Critere, GetValue(GetNomChamp(IndiceChamp)), False);
              Grille.Cells[Col + 1, Row] := RechDom(Critere, GetValue(GetNomChamp(IndiceChamp)), False);
            end
            else
              Grille.Cells[Col, Row] := GetValue(GetNomChamp(IndiceChamp));
          end;
        end;
      end;
  end;
  HMTrad := THSystemMenu(GetControl('HMTrad'));
  if HMTrad <> nil then HMTrad.ResizeGridColumns(Grille);
  if Grille.ColCount > 2 then
    if (Pref = 'PSA') or (Pref = 'PSE') then Grille.col := 3
    else
      if Pref = 'PFO' then Grille.col := 4;
  Grille.Enabled := ZoneAModifier;
  SetControlEnabled('BTNMAJVALUE', True);
  SetControlEnabled('PNMODIF', True);
  SetControlEnabled('BValider', True);
end;
//Affectation de la zone de modication
procedure TOF_PGREAFFECTFORMATION.OnClickBtnMajValue(Sender: TObject);
var
  champ, Libelle, OldValue, NewValue, num: string;
  Col, Row, IndiceChamp, Compteur: integer;
  Obligatoire: Boolean;
begin
  if Grille = nil then exit;
  champ := GetControlText('LISTECHAMP');
  OldValue := GetControlText('OLDVALUE');
  NewValue := GetControlText('NEWVALUE');
  Libelle := '';
  IndiceChamp := -1;
  Compteur := 0;
  if (champ <> '') then
  begin
    Obligatoire := True;
    if champ = 'SO_PGLIBCODESTAT' then Libelle := VH_Paie.PGLibCodeStat
    else
      if Pos('SO_PGLIBORGSTAT', champ) > 0 then
    begin
      num := Copy(Champ, Length(Champ), 1);
      if num = '1' then Libelle := VH_Paie.PGLibelleOrgStat1
      else if num = '2' then Libelle := VH_Paie.PGLibelleOrgStat2
      else if num = '3' then Libelle := VH_Paie.PGLibelleOrgStat3
      else if num = '4' then Libelle := VH_Paie.PGLibelleOrgStat4;
    end
    else
      if Pos('SO_PGLIBCOMBO', champ) > 0 then
    begin
      Obligatoire := False;
      num := Copy(Champ, Length(Champ), 1);
      if num = '1' then Libelle := VH_Paie.PgLibCombo1
      else if num = '2' then Libelle := VH_Paie.PgLibCombo2
      else if num = '3' then Libelle := VH_Paie.PgLibCombo3
      else if num = '4' then Libelle := VH_Paie.PgLibCombo4;
    end
    else
      if Pos('SO_PGFFORLIBRE', champ) > 0 then
    begin
      Obligatoire := False;
      num := Copy(Champ, Length(Champ), 1);
      if num = '1' then Libelle := VH_Paie.FormationLibre1
      else if num = '2' then Libelle := VH_Paie.FormationLibre2
      else if num = '3' then Libelle := VH_Paie.FormationLibre3
      else if num = '4' then Libelle := VH_Paie.FormationLibre4
      else if num = '5' then Libelle := VH_Paie.FormationLibre5
      else if num = '6' then Libelle := VH_Paie.FormationLibre6
      else if num = '7' then Libelle := VH_Paie.FormationLibre7
      else if num = '8' then Libelle := VH_Paie.FormationLibre8;
    end;
    if (Obligatoire) and ((OldValue = '') or (NewValue = '')) then
    begin
      PgiBox(TexteMessage[3], Ecran.caption);
      exit;
    end;
    if Libelle = '' then
    begin
      PgiBox(TexteMessage[4], Ecran.caption);
      exit;
    end;
    for Col := 0 to Grille.ColCount - 1 do
      if Pos(Libelle, Grille.Cells[Col, 0]) > 0 then
      begin
        IndiceChamp := Col;
        Break;
      end;
    if IndiceChamp = -1 then
    begin
      PgiBox(TexteMessage[5], Ecran.caption);
      exit;
    end;
    for Row := 1 to Grille.RowCount - 1 do
    begin
      if (Grille.CellValues[IndiceChamp, Row] = OldValue) and (Grille.CellValues[IndiceChamp + 1, Row] <> NewValue) then
      begin
        Grille.CellValues[IndiceChamp + 1, Row] := NewValue;
        Compteur := Compteur + 1;
        Modifier := True;
      end;
    end;
  end
  else
  begin
    PgiBox(TexteMessage[3], Ecran.caption);
    exit;
  end;
  PgiInfo(IntToStr(Compteur) + ' donnée(s) modifiée(s).', Ecran.Caption);
end;

procedure TOF_PGREAFFECTFORMATION.MiseEnFormeGrille;
begin
  GlbSql := '';
  ZoneAModifier := False;
  Grille.RowCount := 2;
  Grille.ColCount := 4;
  Grille.FixedRows := 1;
  Grille.Titres.Clear;
  TitreColonne.Clear;
  Grille.Titres.Add('Etablissement');
  Grille.ColAligns[0] := tacenter;
  Grille.ColEditables[0] := False;
  Grille.Titres.Add('Salarié');
  Grille.ColAligns[1] := tacenter;
  Grille.ColEditables[1] := False;
  Grille.Titres.Add('Formation');
  Grille.ColFormats[2] := 'CB=PGSTAGEFORM';
  Grille.ColEditables[1] := False;
  Grille.Titres.Add('N° session');
  Grille.ColAligns[1] := tacenter;
  Grille.ColEditables[1] := False;
  TitreColonne.Add(Pref + '_ETABLISSMENT');
  TitreColonne.Add(Pref + '_SALARIE');
  TitreColonne.Add('PFO_CODESTAGE');
  TitreColonne.Add('PFO_ORDRE');
  Grille.ColCount := Grille.ColCount + 2;
  GlbSql := Pref + '_DATEDEBUT,' + Pref + '_DATEFIN,';
  Grille.Titres.Add('Date début');
  Grille.ColAligns[3] := tacenter;
  Grille.ColEditables[2] := False;
  Grille.Titres.Add('Date fin');
  Grille.ColAligns[4] := tacenter;
  Grille.ColEditables[3] := False;
  TitreColonne.Add(Pref + '_DATEDEBUT');
  TitreColonne.Add(Pref + '_DATEFIN');
  if GetControlText('LISTETABLE') = 'LIBRESAL' then
  begin
    MiseEnFormeColGrille('Etablissement', 'ETABLISSEMENT');
    MiseEnFormeColGrille('Libellé emploi', 'LIBEMPLOIFOR');
    MiseEnFormeColGrille('Responsable formation', 'RESPONSFOR');
    MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat1, 'TRAVAILN1');
    MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat2, 'TRAVAILN2');
    MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat3, 'TRAVAILN3');
    MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat4, 'TRAVAILN4');
    MiseEnFormeColGrille(VH_PAIE.PGLibCodeStat, 'CODESTAT');
  end;
  if GetControlText('LISTETABLE') = 'LIBREFORM' then
  begin
    MiseEnFormeColGrille(VH_PAIE.FormationLibre1, 'FORMATION1');
    MiseEnFormeColGrille(VH_PAIE.FormationLibre2, 'FORMATION2');
    MiseEnFormeColGrille(VH_PAIE.FormationLibre3, 'FORMATION3');
    MiseEnFormeColGrille(VH_PAIE.FormationLibre4, 'FORMATION4');
    MiseEnFormeColGrille(VH_PAIE.FormationLibre5, 'FORMATION5');
    MiseEnFormeColGrille(VH_PAIE.FormationLibre6, 'FORMATION6');
    MiseEnFormeColGrille(VH_PAIE.FormationLibre7, 'FORMATION7');
    MiseEnFormeColGrille(VH_PAIE.FormationLibre8, 'FORMATION8');
  end;
  GlbSql := Trim(GlbSql);
  GlbSql := Copy(GlbSql, 1, Length(GlbSql) - 1);
  Grille.UpdateTitres;
end;


procedure TOF_PGREAFFECTFORMATION.MiseEnFormeColGrille(Champ, Suffixe: string);
var
  Tablette: string;
begin
  if Grille = nil then exit;
  if (Champ <> '') and (GetControlText('CK' + Suffixe) = 'X') then
  begin
    ZoneAModifier := True;
    Grille.ColCount := Grille.ColCount + 2;
    Grille.Titres.Add(Champ + ' affecté');
    Grille.Titres.Add(Champ + ' modifié');
    TitreColonne.Add(Pref + '_' + Suffixe);
    TitreColonne.Add(Pref + '_' + Suffixe);
    Grille.CellValues[Grille.ColCount - 2, 0] := Pref + '_' + Suffixe;
    Grille.CellValues[Grille.ColCount - 1, 0] := Pref + '_' + Suffixe;
    if Suffixe = 'ETABLISSEMENT' then Tablette := 'TTETABLISSEMENT'
    else if Suffixe = 'LIBEMPLOIFOR' then tablette := 'PGLIBEMPLOI'
    else if Suffixe = 'RESPONSFOR' then tablette := 'PGSALARIEDEPORTE'
    else Tablette := 'PG' + Suffixe;
    Grille.ColFormats[Grille.ColCount - 2] := 'CB=' + Tablette;
    Grille.ColFormats[Grille.ColCount - 1] := 'CB=' + Tablette;
    Grille.ColEditables[Grille.ColCount - 2] := False;
    GlbSql := GlbSql + Pref + '_' + Suffixe + ',';
  end;
end;

{function TOF_PGREAFFECTFORMATION.RecupTabletteAssocie(Champ: string): string;
begin
  if Champ = '' then result := ''
  else
    if champ = 'SO_PGLIBCODESTAT' then result := 'PGCODESTAT'
  else
    if Pos('SO_PGLIBORGSTAT', champ) > 0 then result := 'PGTRAVAILN' + Copy(Champ, Length(Champ), 1)
  else
    if Pos('SO_PGLIBCOMBO', champ) > 0 then result := 'PGLIBREPCMB' + Copy(Champ, Length(Champ), 1)
  else
    if Pos('SO_PGFFORLIBRE', champ) > 0 then result := 'PGFORMATION' + Copy(Champ, Length(Champ), 1);
end;         }

procedure TOF_PGREAFFECTFORMATION.OnUpdate;
var
  ListeTable: string;
  NbMajOk: integer;
begin
  inherited;
  ListeTable := GetControlText('LISTETABLE');
  NbMajOk := 0;
  if (ListeTable = '') then
  begin
    PGIBox(TexteMessage[1], Ecran.caption);
    exit;
  end;
  TEvent := TStringList.create;
  TEvent.Add(THedit(GetControl('LISTETABLE')).Text);
  if (ListeTable = 'RECOPIEFORM') or (ListeTable = 'RECOPIESAL') then
  begin
    if PgiAsk(TexteMessage[9], Ecran.Caption) = MrNo then exit;
    RecopieTablesSalaries(NbMajOk);
  end
  else
  begin
    if PgiAsk(TexteMessage[10], Ecran.Caption) = MrNo then exit;
    if Modifier or ZoneAModifier then
      MiseAjourSelectif(NbMajOk)
    else
    begin
      PGIBox(TexteMessage[7], Ecran.caption);
      exit;
    end;
  end;
  if NbMajOk = 0 then PGIInfo(TexteMessage[8], Ecran.Caption)
  else
  begin
    PGIInfo(IntToStr(NbMajOk) + ' enregistrement(s) sauvegardé(s).Traitement terminé.', Ecran.Caption);
    if (ListeTable <> 'RECOPIEFORM') and (ListeTable <> 'RECOPIESAL') then OnClickBtnActiveWhere(nil);
  end;
  if TEvent <> nil then TEvent.free;
end;

procedure TOF_PGREAFFECTFORMATION.MiseAjourSelectif(var NbMajOk: integer);
var
  Row, IndiceChamp: integer;
  ListeTable, St, StWhere: string;
begin
  NbMajOk := 0;
  ListeTable := GetControlText('LISTETABLE');
  InitMove(Grille.RowCount, '');
  for Row := 1 to Grille.RowCount - 1 do
    //Recherche par lignes des valeurs modifiées
  begin
    St := ''; //Positionnement sur la prémière colonne de réaffectation
    IndiceChamp := 6;
    RecupClauseChampReaffecter(IndiceChamp, Row, St); // Recherche par Colonnes des valeurs modifiées
    if St <> '' then
    begin
      try
        BeginTrans;
        if ListeTable = 'LIBREFORM' then
        begin
          TEvent.Add('Réaffectation du salarié ' + Grille.Cells[1, Row] + ' pour la formation du ' + GetControlText('PFO_DATEDEBUT') + ' au ' + GetControlText('PFO_DATEFIN'));
          StWhere := ' WHERE ' + Pref + '_SALARIE="' + Grille.Cells[1, Row] + '" AND ' + Pref + '_CODESTAGE="' + Grille.Cells[2, Row] + '" AND ' + Pref + '_ORDRE="' +
            Grille.Cells[3, Row] + '"' +
            'AND PFO_ETABLISSEMENT="' + Grille.Cells[0, Row] + '" ' +
            'AND PFO_DATEFIN>="' + USDateTime(StrToDate(GetControlText('PFO_DATEDEBUT'))) + '" ' +
            'AND PFO_DATEFIN<="' + USDateTime(StrToDate(GetControlText('PFO_DATEFIN'))) + '" ';
          MoveCur(False);
          NbMajOk := NbMajOk + ExecuteSql('UPDATE FORMATIONS SET ' + st + StWhere);
          TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'FORMATIONS', False));
          MoveCur(False);
          CommitTrans;
          TEvent.add('Mise à jour des données OK.');
          Modifier := False;
        end;
        if ListeTable = 'LIBRESAL' then
        begin
          TEvent.Add('Réaffectation du salarié ' + Grille.Cells[1, Row] + ' pour la formation du ' + GetControlText('PFO_DATEDEBUT') + ' au ' + GetControlText('PFO_DATEFIN'));
          StWhere := ' WHERE ' + Pref + '_SALARIE="' + Grille.Cells[1, Row] + '" AND ' + Pref + '_CODESTAGE="' + Grille.Cells[2, Row] + '" AND ' + Pref + '_ORDRE="' +
            Grille.Cells[3, Row] + '"' +
            'AND PFO_ETABLISSEMENT="' + Grille.Cells[0, Row] + '" ' +
            'AND PFO_DATEFIN>="' + USDateTime(StrToDate(GetControlText('PFO_DATEDEBUT'))) + '" ' +
            'AND PFO_DATEFIN<="' + USDateTime(StrToDate(GetControlText('PFO_DATEFIN'))) + '" ';
          MoveCur(False);
          NbMajOk := NbMajOk + ExecuteSql('UPDATE FORMATIONS SET ' + st + StWhere);
          TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'FORMATIONS', False));
          MoveCur(False);
          CommitTrans;
          TEvent.add('Mise à jour des données OK.');
          Modifier := False;
        end;
      except
        Rollback;
        NbMajOk := 0;
        TEvent.add('Une erreur est survenue lors de la mise à jour des données.');
        PGIBox('Une erreur est survenue lors de la mise à jour des données.', Ecran.caption);
      end;
    end;
  end;
  //if NbMajOk>0 then CreeJnalEvt ('004','125','OK',nil,nil,TEvent)
  //else CreeJnalEvt ('004','125','ERR',nil,nil,TEvent);
  FiniMove;
end;

procedure TOF_PGREAFFECTFORMATION.RecopieTablesSalaries(var NbMajOk: integer);
var
  Pages: TPageControl;
  StWhere, Champ: string;
begin
  NbMajOk := 0;
  Champ := '';
  if getControlText('LISTETABLE') = 'RECOPIESAL' then
  begin
    if (GetControlText('CKTRAVAILN1') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN1');
    if (GetControlText('CKTRAVAILN2') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN2');
    if (GetControlText('CKTRAVAILN3') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN3');
    if (GetControlText('CKTRAVAILN4') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN4');
    if (GetControlText('CKCODESTAT') = 'X') then Champ := Champ + RecupSyntaxe('CODESTAT');
    if (GetControlText('CKETABLISSEMENT') = 'X') then Champ := Champ + RecupSyntaxe('ETABLISSEMENT');
    if (GetControlText('CKLIBEMPLOIFOR') = 'X') then Champ := Champ + RecupSyntaxe('LIBEMPLOIFOR');
    if (GetControlText('CKRESPONSFOR') = 'X') then Champ := Champ + RecupSyntaxe('RESPONSFOR');
  end;
  if getControlText('LISTETABLE') = 'RECOPIEFORM' then
  begin
    if (GetControlText('CKFORMATION1') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION1');
    if (GetControlText('CKFORMATION2') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION2');
    if (GetControlText('CKFORMATION3') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION3');
    if (GetControlText('CKFORMATION4') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION4');
    if (GetControlText('CKFORMATION5') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION5');
    if (GetControlText('CKFORMATION6') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION6');
    if (GetControlText('CKFORMATION7') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION7');
    if (GetControlText('CKFORMATION8') = 'X') then Champ := Champ + RecupSyntaxe('FORMATION8');
  end;

  if Champ = '' then
  begin
    PgiBox(TexteMessage[6], Ecran.Caption);
    Exit;
  end;
  InitMove(100, '');
  Pages := TPageControl(GetControl('Pages'));
  SetControlProperty('DATEDEB', 'Name', 'PFO_DATEDEBUT');
  SetControlProperty('DATEFIN', 'Name', 'PFO_DATEFIN');
  StWhere := RecupWhereCritere(PAges);
  TEvent.Add('Réaffectation des formations du ' + GetControlText('PFO_DATEDEBUT') + ' au ' + GetControlText('PFO_DATEFIN'));
  try
    BeginTrans;
    Champ := Copy(Champ, 1, Length(Champ) - 1) + ' ';
    MoveCur(False);
    NbMajOk := NbMajOk + ExecuteSql('UPDATE FORMATIONS SET ' + Champ + StWhere);
    TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'FORMATIONS', False));
    if (GetControlText('CKRESPONSFOR') = 'X') then ExecuteSql('UPDATE FORMATIONS SET PFO_RESPONSFOR="" WHERE PFO_RESPONSFOR IS NULL');
    CommitTrans;
    TEvent.add('Mise à jour OK');
    //  CreeJnalEvt ('004','125','OK',nil,nil,TEvent);
  except
    Rollback;
    NbMajOk := 0;
    TEvent.add('Une erreur est survenue lors de la mise à jour des données.');
    //  CreeJnalEvt ('004','125','ERR',nil,nil,TEvent);
    PGIBox('Une erreur est survenue lors de la mise à jour des données.', Ecran.caption);
  end;
  FiniMove;
end;

function TOF_PGREAFFECTFORMATION.RecupSyntaxe(Champ: string): string;
var
  PrefChamp, ChampSal, Table, jointure: string;
begin
  if Champ <> 'LIBEMPLOIFOR' then ChampSal := Champ
  else ChampSal := 'LIBELLEEMPLOI';
  if getControlText('LISTETABLE') = 'RECOPIEFORM' then
  begin
    PrefChamp := 'PSS';
    Table := 'SESSIONSTAGE';
    Jointure := 'PSS_CODESTAGE=PFO_CODESTAGE AND PSS_MILLESIME=PFO_MILLESIME AND PSS_ORDRE=PFO_ORDRE';
  end
  else
  begin
    If Champ = 'RESPONSFOR' then
    begin
        PrefChamp := 'PSE';
        Table := 'DEPORTSAL';
        Jointure := 'PSE_SALARIE=PFO_SALARIE';
    end
    else
    begin
        PrefChamp := 'PSA';
        Table := 'SALARIES';
        Jointure := 'PSA_SALARIE=PFO_SALARIE';

    end;
  end;
  Result := 'PFO_' + Champ + '=(SELECT ' + PrefChamp + '_' + ChampSal + ' FROM ' + Table + ' WHERE ' + Jointure + '),';
  if Champ = 'TRAVAILN1' then
    TEvent.add('Réaffectation organisation 1')
  else
    if Champ = 'TRAVAILN2' then
    TEvent.add('Réaffectation organisation 2')
  else
    if Champ = 'TRAVAILN3' then
    TEvent.add('Réaffectation organisation 3')
  else
    if Champ = 'TRAVAILN4' then
    TEvent.add('Réaffectation organisation 4')
  else
    if Champ = 'CODESTAT' then
    TEvent.add('Réaffectation code statistique')
  else
    if Champ = 'LIBREPCMB1' then
    TEvent.add('Réaffectation tablette libre 1')
  else
    if Champ = 'LIBREPCMB2' then
    TEvent.add('Réaffectation tablette libre 2')
  else
    if Champ = 'LIBREPCMB3' then
    TEvent.add('Réaffectation tablette libre 3')
  else
    if Champ = 'LIBREPCMB4' then
    TEvent.add('Réaffectation tablette libre 4')
  else
    if Champ = 'FORMATION1' then
    TEvent.add('Réaffectation formation libre 1')
  else
    if Champ = 'FORMATION2' then
    TEvent.add('Réaffectation formation libre 2')
  else
    if Champ = 'FORMATION3' then
    TEvent.add('Réaffectation formation libre 3')
  else
    if Champ = 'FORMATION4' then
    TEvent.add('Réaffectation formation libre 4')
  else
    if Champ = 'FORMATION5' then
    TEvent.add('Réaffectation formation libre 5')
  else
    if Champ = 'FORMATION6' then
    TEvent.add('Réaffectation formation libre 6')
  else
    if Champ = 'FORMATION7' then
    TEvent.add('Réaffectation formation libre 7')
  else
    if Champ = 'FORMATION8' then
    TEvent.add('Réaffectation formation libre 8');
end;

procedure TOF_PGREAFFECTFORMATION.RespElipsisClick(Sender : TObject);
var StFrom,StWhere : String;
begin
	If PGBundleHierarchie Then //PT3
		ElipsisResponsableMultidos(Sender)
	Else
	Begin
        StFrom := 'SALARIES';
        StWhere := 'PSA_SALARIE IN (SELECT PSE_RESPONSFOR FROM DEPORTSAL)';
        LookupList(THEdit(Sender), 'Liste des salariés',StFrom,'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
    End;
end;

procedure TOF_PGREAFFECTFORMATION.ExitEdit(Sender: TObject);//PT2
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;

initialization
  registerclasses([TOF_PGREAFFECTFORMATION]);
end.

