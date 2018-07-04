{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 04/04/2002
Modifié le ... :   /  /
Description .. : Rectification des tables libres
Mots clefs ... : LIBRES;PAIE
*****************************************************************
Structure :
Ce source repose sur une tof Mere TOF_PGREAFFECT qui gère l'évènementiel de la grille nommé GRILLE,
l'affichage des champs et valeurs de prefixe PPU
Les tofs fille restent personnalisés au traitement spécifique de chaque réaffectation (fiche)
PT1  18/10/2002 V585 SB Génération d'évènements pour traçage
PT2-1  25/05/2004 V_50 SB FQ 11055 Réaffectation si champ libre sans valeur initiale
PT2-2  25/05/2004 V_50 SB Refonte gestion tabulation
PT2-3  25/05/2004 V_50 SB Message d'alerte distinct si réaffectation salaries ou historique
PT3    14/01/2005 V_60 SB FQ 11055 Valeur par défaut : domaine d'affectation
PT4    11/08/2005 V_60 PH FQ 12288 Ergonomie Message
PT5    11/05/2007 V_72 JL FQ 14003 Corrections clause maj des absences si champ libres LIBRECMB renseigné dans critères
PT6    28/12/2007 V_80 GGU FQ 14003 aucune réaffectation ne fonctionne ; il manque également les titres de colonne

}
unit UTOFPGReaffectTable;

interface
uses StdCtrls, Controls, Classes, Graphics,  sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF,
  UTOB, HStatus,HQRy, HTB97, Grids, HSysMenu;

type

  TOF_PGREAFFECT = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  private
    Grille: THGrid;
    Tob_Grille: Tob;
    Modifier, ZoneAModifier: Boolean;
    TitreColonne: TStringList;
    TEvent: TStringList; //PT1
    procedure ExitEdit(Sender: TObject);
    procedure GrillePostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    //procedure GridColorCell (ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
    procedure RecupClauseChampReaffecter(IndiceChamp, Row: integer; var st: string);
  end;

  TOF_PGREAFFECTTABLE = class(TOF_PGREAFFECT)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  private
    GlbSql, Pref: string;
    procedure MiseEnFormeGrille;
    procedure MiseEnFormeColGrille(Champ, Suffixe: string);
    procedure OnChangeListeTable(Sender: TObject);
    procedure OnChangeListeChamp(Sender: TObject);
    procedure OnClickBtnActiveWhere(Sender: TObject);
    procedure OnClickBtnMajValue(Sender: TObject);
    function RecupTabletteAssocie(Champ: string): string;
    //     Function  RecupNomTableAssocie (Champ : string) : String;
    procedure RecopieTablesSalaries(var NbMajOk: integer);
    procedure MiseAjourSelectif(var NbMajOk: integer);
    function RendClauseAbsenceSalarie(StClause: string): string; //PT5
    function RendWhereAbsenceSalarie(StClause: string): string;
    function RecupSyntaxe(Champ: string): string;
  end;

  TOF_PGREAFFECTORGANISME = class(TOF_PGREAFFECT)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  private
    procedure MiseEnFormeGrille;
    procedure OnClickBtnActiveWhere(Sender: TObject);
    procedure OnClickBtnMajValue(Sender: TObject);
    procedure OnClickHisto(Sender: TObject);
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
    {8}'Aucune donnée n''a été réaffectée.', // PT4
    {9}'Attention! Vous demandez à recopier les champs libres de la table salarié#13#10 sur les champs libres de l''historique de paie.#13#10 Voulez-vous continuer le traitement?',
    {10}'Attention! Vous demandez à mettre à jour les champs libres de l''historique de paie.#13#10 Voulez-vous continuer le traitement?',
    {11}'Veuillez cliquer sur le bouton de validation vert pour la recopie des données.',
    {12}'Attention! Vous demandez à mettre à jour les champs libres de la table salarié.#13#10 Voulez-vous continuer le traitement?' { PT2-3 }
    );

implementation
uses EntPaie, PGEditOutils, PGEditOutils2, P5Util, Pgoutils,PgOutils2, PGEdtEtat, P5Def;

{ TOF_PGREAFFECT }

procedure TOF_PGREAFFECT.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if edit.text <> '' then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
end;

{procedure TOF_PGREAFFECT.GridColorCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
if not ZoneAModifier then exit;
if Pos('modifié',Grille.cells[ACol,0])>=1 then
   If Grille.Cells[Acol,Arow]<>Grille.Cells[Acol-1,Arow] then
     if (Grille.Cells[Acol,Arow]<>'') AND (Grille.CellValues[Acol,Arow]<>'NON') then
       //Grille.Cells[Acol,Arow] Canvas.Pen.Color  := clRed ;
end;}

procedure TOF_PGREAFFECT.GrilleCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if not ZoneAModifier then exit;
  if Pos('modifié', Grille.cells[Grille.col, 0]) < 1 then
  begin                         { DEB PT2-2 }
    { Tabulations sens nornmal }
    if (Acol < Grille.col ) or (Arow < Grille.row) then
      Begin
      if (Grille.col = Grille.ColCount - 1) then
        begin
        Grille.col := 0;
        if (Grille.Row <> Grille.RowCount - 1) then Grille.Row := Grille.Row + 1;
        end
      else
        Grille.Col := Grille.Col + 1;
      End
    else
    { Tabulations sens inverse }
      if (Acol > Grille.col ) then
        Begin
        if (Grille.col = 0) then
          Begin
          Grille.col := Grille.colCount - 1;
          if (Grille.Row > 1 ) then  Grille.Row := Grille.Row - 1
          else                       Grille.Row := Grille.RowCount - 1;
          End
        else
          Grille.Col := Grille.Col - 1;
        End;                   { FIN PT2-2 }
  end;
end;

procedure TOF_PGREAFFECT.GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if not ZoneAModifier then exit;
  if Pos('modifié', Grille.Cells[ACol, 0]) < 1 then
    GridGriseCell(Grille, Acol, Arow, Canvas);
end;

procedure TOF_PGREAFFECT.OnArgument(Arguments: string);
var
  Min, Max: string;
  Edit: ThEdit;
begin
  inherited;
  Modifier := False;
  ZoneAModifier := False;
  VisibiliteChamp(Ecran);
  VisibiliteChampLibre(Ecran);
  //Valeur par défaut
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Edit := ThEdit(getcontrol('PPU_SALARIE'));
  if Edit <> nil then
  begin
    Edit.text := Min;
    Edit.OnExit := ExitEdit;
  end;
  Edit := ThEdit(getcontrol('PPU_SALARIE_'));
  if Edit <> nil then
  begin
    Edit.text := Max;
    Edit.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Edit := ThEdit(getcontrol('PPU_ETABLISSEMENT'));
  if Edit <> nil then
  begin
    Edit.text := Min;
  end;
  Edit := ThEdit(getcontrol('PPU_ETABLISSEMENT_'));
  if Edit <> nil then
  begin
    Edit.text := Max;
  end;
  Grille := THGrid(GetControl('GRILLE'));
  if Grille <> nil then
  begin
    Grille.PostDrawCell := GrillePostDrawCell;
    Grille.OnCellEnter := GrilleCellEnter;
    // Grille.GetCellCanvas := GridColorCell ;
  end;
  TitreColonne := TStringList.Create;

  SetControlVisible('CKTRAVAILN1', (VH_PAIE.PGLibelleOrgStat1 <> ''));
  SetControlVisible('CKTRAVAILN2', (VH_PAIE.PGLibelleOrgStat2 <> ''));
  SetControlVisible('CKTRAVAILN3', (VH_PAIE.PGLibelleOrgStat3 <> ''));
  SetControlVisible('CKTRAVAILN4', (VH_PAIE.PGLibelleOrgStat4 <> ''));
  SetControlVisible('CKCODESTAT', (VH_PAIE.PGLibCodeStat <> ''));
  SetControlVisible('CKLIBREPCMB1', (VH_PAIE.PgLibCombo1 <> ''));
  SetControlVisible('CKLIBREPCMB2', (VH_PAIE.PgLibCombo2 <> ''));
  SetControlVisible('CKLIBREPCMB3', (VH_PAIE.PgLibCombo3 <> ''));
  SetControlVisible('CKLIBREPCMB4', (VH_PAIE.PgLibCombo4 <> ''));
end;



procedure TOF_PGREAFFECT.OnClose;
begin
  inherited;
  if TitreColonne <> nil then TitreColonne.Free;
  if Tob_Grille <> nil then Tob_Grille.free;
end;

procedure TOF_PGREAFFECT.RecupClauseChampReaffecter(IndiceChamp, Row: integer; var st: string);
begin
  St := '';
  if IndiceChamp = -1 then exit;
  repeat // Recherche par Colonnes des valeurs modifiées
    if (Grille.CellValues[IndiceChamp, Row]) <> (Grille.CellValues[IndiceChamp + 1, Row]) then
    begin
      if St <> '' then St := St + ',';
      St := St + TitreColonne.Strings[IndiceChamp] + '="' + Grille.CellValues[IndiceChamp + 1, Row] + '" ';
      TEvent.Add(Grille.Titres.Strings[IndiceChamp + 1] + ' ancienne valeur : ' + Grille.CellValues[IndiceChamp, Row] + ' nouvelle valeur : ' + Grille.CellValues[IndiceChamp + 1,
        Row]); //PT1
    end;
    IndiceChamp := IndiceChamp + 2;
  until IndiceChamp >= Grille.ColCount - 1;
end;

{ TOF_PGREAFFECTTABLE }



procedure TOF_PGREAFFECTTABLE.OnArgument(Arguments: string);
var
  Combo: THValComboBox;
  Btn: TToolBarButton97;
begin
  inherited;
  //Gestionnnaire d'evenement
  Combo := THValComboBox(GetControl('LISTETABLE'));
  if Combo <> nil then
    Begin
    Combo.OnChange := OnChangeListeTable;
    Combo.Itemindex := 0;      { PT3 }
    OnChangeListeTable(Combo); { PT3 }
    End;
  Combo := THValComboBox(GetControl('LISTECHAMP'));
  if Combo <> nil then Combo.OnChange := OnChangeListeChamp;
  Btn := TToolBarButton97(GetControl('BTNACTIVEWHERE'));
  if btn <> nil then Btn.Onclick := OnClickBtnActiveWhere;
  Btn := TToolBarButton97(GetControl('BTNMAJVALUE'));
  if btn <> nil then Btn.Onclick := OnClickBtnMajValue;
end;



procedure TOF_PGREAFFECTTABLE.OnChangeListeTable(Sender: TObject);
begin
  SetControlVisible('PANEL', (GetControlText('LISTETABLE') <> 'RECOPIE'));
  SetControlVisible('DATEDEB', (GetControlText('LISTETABLE') <> 'SALARIES'));
  SetControlVisible('DATEFIN', (GetControlText('LISTETABLE') <> 'SALARIES'));
  SetControlVisible('TDATEDEB', (GetControlText('LISTETABLE') <> 'SALARIES'));
  SetControlVisible('TDATEFIN', (GetControlText('LISTETABLE') <> 'SALARIES'));
  SetControlEnabled('BValider', (GetControlText('LISTETABLE') = 'RECOPIE'));
end;

procedure TOF_PGREAFFECTTABLE.OnChangeListeChamp(Sender: TObject);
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
end;


procedure TOF_PGREAFFECTTABLE.OnClickBtnActiveWhere(Sender: TObject);
var
  ListeTable, StSql, ChSql, Critere, AutrePref: string;
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
  if (ListeTable = 'RECOPIE') then
  begin
    PGIBox(TexteMessage[11], Ecran.caption);
    exit;
  end;
  if Modifier then
    if PGIAsk(TexteMessage[2], Ecran.caption) = MrNo then
      exit;
  Pref := TableToPrefixe(ListeTable);
  if Pref = 'PSA' then AutrePref := 'PPU' else AutrePref := 'PSA';
  MiseEnFormeGrille;
  Pages := TPageControl(GetControl('Pages'));
  if Pages <> nil then Critere := RecupWhereCritere(Pages)
  else Critere := '';
  if (GlbSql <> '') then GlbSql := ConvertPrefixe(GlbSql, AutrePref, Pref);
  if Critere <> '' then
  begin
    Critere := ConvertPrefixe(Critere, AutrePref, Pref);
    if Pref = 'PPU' then Critere := Critere + 'AND PPU_DATEFIN>="' + USDateTime(StrToDate(GetControlText('DATEDEB'))) + '" ' +
      'AND PPU_DATEFIN<="' + USDateTime(StrToDate(GetControlText('DATEFIN'))) + '"';
  end
  else
    if Pref = 'PPU' then Critere := 'WHERE PPU_DATEFIN>="' + USDateTime(StrToDate(GetControlText('DATEDEB'))) + '" ' +
    'AND PPU_DATEFIN<="' + USDateTime(StrToDate(GetControlText('DATEFIN'))) + '"';
  if GlbSql <> '' then
    ChSql := Pref + '_ETABLISSEMENT,' + Pref + '_SALARIE,' + GlbSql
  else
    ChSql := Pref + '_ETABLISSEMENT,' + Pref + '_SALARIE ';
  StSql := 'SELECT DISTINCT ' + ChSql + ' FROM ' + ListeTable + ' ' + Critere + ' ORDER BY ' + PRef + '_SALARIE';
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
    if Pref = 'PSA' then Grille.col := 3
    else
      if Pref = 'PPU' then Grille.col := 4;
  Grille.Enabled := ZoneAModifier;
  SetControlEnabled('BTNMAJVALUE', True);
  SetControlEnabled('PNMODIF', True);
  SetControlEnabled('BValider', True);
end;
//Affectation de la zone de modication
procedure TOF_PGREAFFECTTABLE.OnClickBtnMajValue(Sender: TObject);
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
    end;
    if (Obligatoire) and (NewValue = '') then { PT2-1 Suppr Clause. }
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
        //    if (Grille.Cells[IndiceChamp,Row]=RechDom(RecupTabletteAssocie(champ),OldValue,False)) and (Grille.Cells[IndiceChamp+1,Row]<>RechDom(RecupTabletteAssocie(champ),NewValue,False)) then
      begin
        Grille.CellValues[IndiceChamp + 1, Row] := NewValue;
        //      RechDom(RecupTabletteAssocie(champ),NewValue,False);
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

procedure TOF_PGREAFFECTTABLE.MiseEnFormeGrille;
begin
  GlbSql := '';
  ZoneAModifier := False;
  Grille.RowCount := 2;
  Grille.ColCount := 2;
  Grille.FixedRows := 1;
  Grille.Titres.Clear;
  TitreColonne.Clear;
  Grille.Titres.Add('Etablissement');
  Grille.Cells[0,0] := 'Etablissement';  //PT6
  Grille.ColAligns[0] := tacenter;
  Grille.ColEditables[0] := False;
  TitreColonne.Add(Pref + '_ETABLISSMENT');
  Grille.Titres.Add('Salarié');
  Grille.Cells[1,0] := 'Salarié';  //PT6
  Grille.ColAligns[1] := tacenter;
  Grille.ColEditables[1] := False;
  TitreColonne.Add(Pref + '_SALARIE');
  if GetControlText('LISTETABLE') = 'PAIEENCOURS' then
  begin
    Grille.ColCount := Grille.ColCount + 2;
    GlbSql := Pref + '_DATEDEBUT,' + Pref + '_DATEFIN,';
    Grille.Titres.Add('Paie du');
    Grille.Cells[2,0] := 'Salarié';  //PT6
    Grille.ColAligns[2] := tacenter;
    Grille.ColEditables[2] := False;
    TitreColonne.Add(Pref + '_DATEDEBUT');
    Grille.Titres.Add('Paie au');
    Grille.Cells[3,0] := 'Salarié';  //PT6
    Grille.ColAligns[3] := tacenter;
    Grille.ColEditables[3] := False;
    TitreColonne.Add(Pref + '_DATEFIN');
  end;
  MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat1, 'TRAVAILN1');
  MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat2, 'TRAVAILN2');
  MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat3, 'TRAVAILN3');
  MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat4, 'TRAVAILN4');
  MiseEnFormeColGrille(VH_PAIE.PGLibCodeStat, 'CODESTAT');
  MiseEnFormeColGrille(VH_PAIE.PgLibCombo1, 'LIBREPCMB1');
  MiseEnFormeColGrille(VH_PAIE.PgLibCombo2, 'LIBREPCMB2');
  MiseEnFormeColGrille(VH_PAIE.PgLibCombo3, 'LIBREPCMB3');
  MiseEnFormeColGrille(VH_PAIE.PgLibCombo4, 'LIBREPCMB4');

  GlbSql := Trim(GlbSql);
  GlbSql := Copy(GlbSql, 1, Length(GlbSql) - 1);
  Grille.UpdateTitres;
end;


procedure TOF_PGREAFFECTTABLE.MiseEnFormeColGrille(Champ, Suffixe: string);
begin
  if Grille = nil then exit;
  if (Champ <> '') and (GetControlText('CK' + Suffixe) = 'X') then
  begin
    ZoneAModifier := True;
    Grille.ColCount := Grille.ColCount + 2;
    Grille.Titres.Add(Champ + ' affecté');
    TitreColonne.Add(Pref + '_' + Suffixe);
    Grille.Cells[Grille.ColCount - 2,0] := Champ + ' affecté';  //PT6
//PT6    Grille.CellValues[Grille.ColCount - 2, 0] := Pref + '_' + Suffixe;
    Grille.ColFormats[Grille.ColCount - 2] := 'CB=PG' + Suffixe;
    Grille.ColEditables[Grille.ColCount - 2] := False;
    Grille.Titres.Add(Champ + ' modifié');
    Grille.Cells[Grille.ColCount - 1,0] := Champ + ' modifié';  //PT6
//PT6    Grille.CellValues[Grille.ColCount - 1, 0] := Pref + '_' + Suffixe;
    Grille.ColFormats[Grille.ColCount - 1] := 'CB=PG' + Suffixe;
    Grille.ColEditables[Grille.ColCount - 1] := True;
    TitreColonne.Add(Pref + '_' + Suffixe);
    GlbSql := GlbSql + Pref + '_' + Suffixe + ',';
  end;
end;

function TOF_PGREAFFECTTABLE.RecupTabletteAssocie(Champ: string): string;
begin
  if Champ = '' then result := ''
  else
    if champ = 'SO_PGLIBCODESTAT' then result := 'PGCODESTAT'
  else
    if Pos('SO_PGLIBORGSTAT', champ) > 0 then result := 'PGTRAVAILN' + Copy(Champ, Length(Champ), 1)
  else
    if Pos('SO_PGLIBCOMBO', champ) > 0 then result := 'PGLIBREPCMB' + Copy(Champ, Length(Champ), 1);
end;

{function TOF_PGREAFFECTTABLE.RecupNomTableAssocie(Champ: string): String;
begin
if Champ='' then result:=''
 else
  if champ='SO_PGLIBCODESTAT' then result:=Pref+'_CODESTAT'
  else
    if Pos('SO_PGLIBORGSTAT',champ)>0 then result:=Pref+'_TRAVAILN'+Copy(Champ,Length(Champ),1)
    else
      if Pos('SO_PGLIBCOMBO',champ)>0 then result:=Pref+'_LIBREPCMB'+Copy(Champ,Length(Champ),1);
end;                                             }


procedure TOF_PGREAFFECTTABLE.OnUpdate;
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
  TEvent := TStringList.create; //PT1
  TEvent.Add(THedit(GetControl('LISTETABLE')).Text); //PT1
  if (ListeTable = 'RECOPIE') then
  begin
    if PgiAsk(TexteMessage[9], Ecran.Caption) = MrNo then exit;
    RecopieTablesSalaries(NbMajOk);
  end
  else
  begin       { DEB PT2-3 }
    if (ListeTable = 'PAIEENCOURS') then
       if PgiAsk(TexteMessage[10], Ecran.Caption) = MrNo then exit;
    if (ListeTable = 'SALARIES') then
       if PgiAsk(TexteMessage[12], Ecran.Caption) = MrNo then exit;
    if Modifier or ZoneAModifier then
      MiseAjourSelectif(NbMajOk)
    else      { FIN PT2-3 }  
    begin
      PGIBox(TexteMessage[7], Ecran.caption);
      exit;
    end;
  end;
  if NbMajOk = 0 then PGIInfo(TexteMessage[8], Ecran.Caption)
  else
  begin
    PGIInfo(IntToStr(NbMajOk) + ' enregistrement(s) sauvegardé(s).Traitement terminé.', Ecran.Caption);
    if (ListeTable <> 'RECOPIE') then OnClickBtnActiveWhere(nil);
  end;
  if TEvent <> nil then TEvent.free; //PT1
end;

procedure TOF_PGREAFFECTTABLE.MiseAjourSelectif(var NbMajOk: integer);
var
  Row, IndiceChamp: integer;
  ListeTable, St, StWhere: string;
begin
  NbMajOk := 0;
  IndiceChamp := -1;
  ListeTable := GetControlText('LISTETABLE');
  InitMove(Grille.RowCount, '');
  for Row := 1 to Grille.RowCount - 1 do
    //Recherche par lignes des valeurs modifiées
  begin
    St := ''; //Positionnement sur la prémière colonne de réaffectation
    if ListeTable = 'SALARIES' then IndiceChamp := 2
    else
      if ListeTable = 'PAIEENCOURS' then IndiceChamp := 4
    else Break;
    RecupClauseChampReaffecter(IndiceChamp, Row, St); // Recherche par Colonnes des valeurs modifiées
    if St <> '' then
    begin
      try
        BeginTrans;
        if ListeTable = 'PAIEENCOURS' then
        begin
          TEvent.Add('Réaffectation du salarié ' + Grille.Cells[1, Row] + ' sur la période de paie du ' + GetControlText('DATEDEB') + ' au ' + GetControlText('DATEFIN')); //PT1
          StWhere := ' WHERE ' + Pref + '_SALARIE="' + Grille.Cells[1, Row] + '" ' +
            'AND ' + Pref + '_ETABLISSEMENT="' + Grille.Cells[0, Row] + '" ' +
            'AND ' + Pref + '_DATEFIN>="' + USDateTime(StrToDate(GetControlText('DATEDEB'))) + '" ' +
            'AND ' + Pref + '_DATEFIN<="' + USDateTime(StrToDate(GetControlText('DATEFIN'))) + '" ';
          MoveCur(False);
          NbMajOk := NbMajOk + ExecuteSql('UPDATE ' + LISTETABLE + ' SET ' + st + StWhere);
          TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'PAIEENCOURS', False)); //PT1
          St := ConvertPrefixe(St, 'PPU', 'PHB');
          StWhere := ConvertPrefixe(StWhere, 'PPU', 'PHB');
          MoveCur(False);
          NbMajOk := NbMajOk + ExecuteSql('UPDATE HISTOBULLETIN SET ' + st + ' ' + StWhere);
          TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'HISTOBULLETIN', False)); //PT1
          St := ConvertPrefixe(St, 'PHB', 'PHC');
          StWhere := ConvertPrefixe(StWhere, 'PHB', 'PHC');
          MoveCur(False);
          NbMajOk := NbMajOk + ExecuteSql('UPDATE HISTOCUMSAL SET ' + st + ' ' + StWhere);
          TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'HISTOCUMSAL', False)); //PT1
          St := ConvertPrefixe(St, 'PHC', 'PHA');
          StWhere := ConvertPrefixe(StWhere, 'PHC', 'PHA');
          MoveCur(False);
          NbMajOk := NbMajOk + ExecuteSql('UPDATE HISTOANALPAIE SET ' + st + ' ' + StWhere);
          TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'HISTOANALPAIE', False)); //PT1
          St := ConvertPrefixe(St, 'PHA', 'PCN');
          StWhere := ConvertPrefixe(StWhere, 'PHA', 'PCN');
          MoveCur(False);
          St := RendClauseAbsenceSalarie(St); 
          if St <> '' then
          begin
            NbMajOk := NbMajOk + ExecuteSql('UPDATE ABSENCESALARIE SET ' + St + ' ' + StWhere);
            TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'ABSENCESALARIE', False)); //PT1
          end;
        end
        else
        begin
          St := 'UPDATE ' + LISTETABLE + ' SET ' + St +
            ' WHERE ' + Pref + '_SALARIE="' + Grille.Cells[1, Row] + '" AND ' + Pref + '_ETABLISSEMENT="' + Grille.Cells[0, Row] + '" ';
          NbMajOk := NbMajOk + ExecuteSql(st);
          TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', LISTETABLE, False)); //PT1
        end;
        MoveCur(False);
        CommitTrans;
        TEvent.add('Mise à jour des données OK.'); //PT1
        Modifier := False;
      except
        Rollback;
        NbMajOk := 0;
        TEvent.add('Une erreur est survenue lors de la mise à jour des données.'); //PT1
        PGIBox('Une erreur est survenue lors de la mise à jour des données.', Ecran.caption);
      end;
    end;
  end;
  if NbMajOk > 0 then CreeJnalEvt('004', '125', 'OK', nil, nil, TEvent) //PT1
  else CreeJnalEvt('004', '125', 'ERR', nil, nil, TEvent); //PT1
  FiniMove;
end;

procedure TOF_PGREAFFECTTABLE.RecopieTablesSalaries(var NbMajOk: integer);
var
  Pages: TPageControl;
  StWhere, Champ: string;
begin
  NbMajOk := 0;
  Champ := '';
  if (GetControlText('CKTRAVAILN1') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN1');
  if (GetControlText('CKTRAVAILN2') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN2');
  if (GetControlText('CKTRAVAILN3') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN3');
  if (GetControlText('CKTRAVAILN4') = 'X') then Champ := Champ + RecupSyntaxe('TRAVAILN4');
  if (GetControlText('CKCODESTAT') = 'X') then Champ := Champ + RecupSyntaxe('CODESTAT');
  if (GetControlText('CKLIBREPCMB1') = 'X') then Champ := Champ + RecupSyntaxe('LIBREPCMB1');
  if (GetControlText('CKLIBREPCMB2') = 'X') then Champ := Champ + RecupSyntaxe('LIBREPCMB2');
  if (GetControlText('CKLIBREPCMB3') = 'X') then Champ := Champ + RecupSyntaxe('LIBREPCMB3');
  if (GetControlText('CKLIBREPCMB4') = 'X') then Champ := Champ + RecupSyntaxe('LIBREPCMB4');

  if Champ = '' then
  begin
    PgiBox(TexteMessage[6], Ecran.Caption);
    Exit;
  end;
  InitMove(100, '');
  Pages := TPageControl(GetControl('Pages'));
  SetControlProperty('DATEDEB', 'Name', 'PPU_DATEDEBUT');
  SetControlProperty('DATEFIN', 'Name', 'PPU_DATEFIN');
  StWhere := RecupWhereCritere(PAges);
  SetControlProperty('PPU_DATEDEBUT', 'Name', 'DATEDEB');
  SetControlProperty('PPU_DATEFIN', 'Name', 'DATEFIN');
  TEvent.Add('Réaffectation sur la période de paie du ' + GetControlText('DATEDEB') + ' au ' + GetControlText('DATEFIN')); //PT1
  try
    BeginTrans;
    Champ := Copy(Champ, 1, Length(Champ) - 1) + ' ';
    MoveCur(False);
    NbMajOk := NbMajOk + ExecuteSql('UPDATE PAIEENCOURS SET ' + Champ + StWhere);
    TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'PAIEENCOURS', False)); //PT1
    Champ := ConvertPrefixe(Champ, 'PPU', 'PHB');
    StWhere := ConvertPrefixe(StWhere, 'PPU', 'PHB');
    MoveCur(False);
    NbMajOk := NbMajOk + ExecuteSql('UPDATE HISTOBULLETIN SET ' + Champ + ' ' + StWhere);
    TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'HISTOBULLETIN', False)); //PT1
    Champ := ConvertPrefixe(Champ, 'PHB', 'PHC');
    StWhere := ConvertPrefixe(StWhere, 'PHB', 'PHC');
    MoveCur(False);
    NbMajOk := NbMajOk + ExecuteSql('UPDATE HISTOCUMSAL SET ' + Champ + ' ' + StWhere);
    TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'HISTOCUMSAL', False)); //PT1
    Champ := ConvertPrefixe(Champ, 'PHC', 'PHA');
    StWhere := ConvertPrefixe(StWhere, 'PHC', 'PHA');
    MoveCur(False);
    NbMajOk := NbMajOk + ExecuteSql('UPDATE HISTOANALPAIE SET ' + Champ + ' ' + StWhere);
    TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'HISTOANALPAIE', False)); //PT1
    Champ := ConvertPrefixe(Champ, 'PHA', 'PCN');
    StWhere := ConvertPrefixe(StWhere, 'PHA', 'PCN');
    StWhere := RendWhereAbsenceSalarie(StWhere);    //PT5
    MoveCur(False);
    Champ := RendClauseAbsenceSalarie(Champ);
    if Champ <> '' then
    begin
      NbMajOk := NbMajOk + ExecuteSql('UPDATE ABSENCESALARIE SET ' + Champ + ' ' + StWhere);
      TEvent.add('Traitement de la table : ' + RechDom('TTTABLES', 'ABSENCESALARIE', False)); //PT1
    end;
    MoveCur(False);
    CommitTrans;
    TEvent.add('Mise à jour OK'); //PT1
    CreeJnalEvt('004', '125', 'OK', nil, nil, TEvent); //PT1
  except
    Rollback;
    NbMajOk := 0;
    TEvent.add('Une erreur est survenue lors de la mise à jour des données.');
    CreeJnalEvt('004', '125', 'ERR', nil, nil, TEvent); //PT1
    PGIBox('Une erreur est survenue lors de la mise à jour des données.', Ecran.caption);
  end;
  FiniMove;
end;


function TOF_PGREAFFECTTABLE.RendWhereAbsenceSalarie(StClause: string): string; //PT5
var
  Clause: string;
  longueur : Integer;
begin
  result := '';
  while (StClause <> '') do
  begin
    Clause := ReadTokenPipe(StClause, ' AND ');
    if Pos('LIBREPCMB', Clause) < 1 then
    begin
      if result <> '' then result := result + ' AND ';
      result := result + Clause;
    end;
  end;
  If Result <> '' then
  begin
    longueur := Length(Result);
    If Copy(Result,Longueur,1) <> ')' then Result := Result + ')';
  end;
end;


function TOF_PGREAFFECTTABLE.RendClauseAbsenceSalarie(StClause: string): string;
var
  Clause: string;
begin
  result := '';
  while (StClause <> '') do
  begin
    Clause := ReadTokenPipe(StClause, ',');
    if Pos('LIBREPCMB', Clause) < 1 then
    begin
      if result <> '' then result := result + ',';
      result := result + Clause;
    end;
  end;
end;
//DEB PT1 ajout function pour traitement evenement
function TOF_PGREAFFECTTABLE.RecupSyntaxe(Champ: string): string;
begin
  Result := 'PPU_' + Champ + '=(SELECT PSA_' + Champ + ' FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
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
    TEvent.add('Réaffectation tablette libre 4');
end;
//FIN
{ TOF_PGREAFFECTORGANISME }

procedure TOF_PGREAFFECTORGANISME.MiseEnFormeGrille;
begin
  TitreColonne.Clear;
  Grille.ColCount := 6;
  TitreColonne.Add('PCT_RUBRIQUE');
  Grille.Titres.Add('Rubrique');
  Grille.ColAligns[0] := tacenter;
  Grille.ColEditables[0] := True;
  Grille.Cells[0,0] := 'Rubrique';  //PT6
//PT6  Grille.CellValues[0, 0] := 'PCT_RUBRIQUE';
  TitreColonne.Add('PCT_LIBELLE');
  Grille.Titres.Add('Libellé');
  Grille.ColAligns[1] := tacenter;
  Grille.ColEditables[1] := True;
  Grille.Cells[1,0] := 'Libellé';  //PT6
//PT6  Grille.CellValues[1, 0] := 'PCT_LIBELLE';
  TitreColonne.Add('PCT_PREDEFINI');
  Grille.Titres.Add('Prédefini');
  Grille.ColAligns[2] := tacenter;
  Grille.ColEditables[2] := True;
  Grille.ColFormats[2] := 'CB=PGPREDEFINI';
  Grille.Cells[2,0] := 'Prédefini';  //PT6
//PT6  Grille.CellValues[2, 0] := 'PCT_PREDEFINI';
  TitreColonne.Add('PCT_ORGANISME');
  Grille.Titres.Add('Organisme affecté');
  Grille.ColAligns[3] := tacenter;
  Grille.ColEditables[3] := True;
  Grille.ColFormats[3] := 'CB=PGTYPEORGANISME';
  Grille.Cells[3,0] := 'Organisme affecté';  //PT6
//PT6  Grille.CellValues[3, 0] := 'PCT_ORGANISME';
  TitreColonne.Add('PCT_ORGANISME');
  Grille.Titres.Add('Organisme modifié');
  Grille.ColAligns[4] := tacenter;
  Grille.ColEditables[4] := True;
  Grille.ColFormats[4] := 'CB=PGTYPEORGANISME';
  Grille.Cells[4,0] := 'Organisme modifié';  //PT6
//PT6  Grille.CellValues[4, 0] := 'PCT_ORGANISME';
  TitreColonne.Add('HISTO');
  Grille.Titres.Add('modifié l''historique');
  Grille.ColAligns[5] := tacenter;
  Grille.ColEditables[5] := True;
  Grille.ColFormats[5] := 'CB=PGOUINON'; // IntToStr(Ord(csCheckbox)); //csCoche
  Grille.Cells[5,0] := 'modifié l''historique';  //PT6
//PT6  Grille.CellValues[5, 0] := 'HISTO';
  //Grille.ColTypes[5]:='B';
  ZoneAModifier := True;
end;

procedure TOF_PGREAFFECTORGANISME.OnArgument(Arguments: string);
var
  Btn: TToolBarButton97;
  Check: TCheckBox;
begin
  inherited;
  //Gestionnnaire d'evenement
  Btn := TToolBarButton97(GetControl('BTNACTIVEWHERE'));
  if btn <> nil then Btn.Onclick := OnClickBtnActiveWhere;
  Btn := TToolBarButton97(GetControl('BTNMAJVALUE'));
  if btn <> nil then Btn.Onclick := OnClickBtnMajValue;
  Check := TCheckBox(GetControl('CKHISTO'));
  if Check <> nil then Check.OnClick := OnClickHisto;
end;

procedure TOF_PGREAFFECTORGANISME.OnClickHisto(Sender: TObject);
var
  i: integer;
begin
  SetControlEnabled('TDATEDEB', GetControlText('CKHISTO') = 'X');
  SetControlEnabled('DATEDEB', GetControlText('CKHISTO') = 'X');
  SetControlEnabled('TDATEFIN', GetControlText('CKHISTO') = 'X');
  SetControlEnabled('DATEFIN', GetControlText('CKHISTO') = 'X');
  for i := 1 to Grille.RowCount - 1 do
    if GetControlText('CKHISTO') = 'X' then
      Grille.CellValues[5, i] := 'OUI'
    else
      Grille.CellValues[5, i] := 'NON';


end;

procedure TOF_PGREAFFECTORGANISME.OnClickBtnActiveWhere(Sender: TObject);
var
  StSql, Critere: string;
  Q: TQuery;
  Pages: TPageControl;
  i: integer;
  HMTrad: THSystemMenu;
  CEG, STD, DOS: Boolean;
begin
  Grille.RowCount := 2;
  Grille.ColCount := 2;
  MiseEnFormeGrille;
  Pages := TPageControl(GetControl('Pages'));
  if Pages <> nil then Critere := RecupWhereCritere(Pages)
  else Critere := '';
  if Critere <> '' then Critere := Critere + ' AND'
  else Critere := Critere + ' WHERE';
  AccesPredefini('TOUS', CEG, STD, DOS);
  if not CEG then
    if STD then Critere := Critere + ' PCT_PREDEFINI<>"CEG" AND'
    else
      if DOS then Critere := Critere + ' PCT_PREDEFINI="DOS" AND';
  StSql := 'SELECT PCT_RUBRIQUE,PCT_LIBELLE,PCT_PREDEFINI,PCT_ORGANISME FROM COTISATION  ' +
    Critere + ' ##PCT_PREDEFINI## PCT_NATURERUB="COT" ORDER BY PCT_RUBRIQUE';
  Q := OpenSql(StSql, True);
  if Tob_Grille <> nil then
  begin
    Tob_Grille.free;
    Tob_Grille := nil;
  end;
  Tob_Grille := Tob.create('La liste à maj', nil, -1);
  Tob_Grille.LoadDetailDB('La liste à maj', '', '', Q, False);
  for i := 0 to Tob_Grille.detail.count - 1 do
  begin
    Tob_Grille.Detail[i].AddChampSupValeur('PCT_ORGANISME_', Tob_Grille.Detail[i].GetValue('PCT_ORGANISME'));
    Tob_Grille.Detail[i].AddChampSupValeur('HISTO', 'OUI');
  end;
  Ferme(Q);
  Tob_Grille.PutGridDetail(Grille, False, False, '');
  MiseEnFormeGrille;
  Grille.Enabled := ZoneAModifier;
  Grille.UpdateTitres;
  HMTrad := THSystemMenu(GetControl('HMTrad'));
  if HMTrad <> nil then HMTrad.ResizeGridColumns(Grille);
  Grille.Col := 5;
  SetControlEnabled('BTNMAJVALUE', True);
  SetControlEnabled('PNMODIF', True);
  SetControlEnabled('BValider', True);
end;

procedure TOF_PGREAFFECTORGANISME.OnClickBtnMajValue(Sender: TObject);
var
  Row, Compteur: integer;
  OldValue, NewValue: string;
begin
  OldValue := GetControlText('OLDVALUE');
  NewValue := GetControlText('NEWVALUE');
  Compteur := 0;
  if (OldValue <> '') and (NewValue <> '') then
  begin
    for Row := 1 to Grille.RowCount - 1 do
    begin
      if (Grille.CellValues[3, Row] = OldValue) and (Grille.CellValues[4, Row] <> NewValue) then
      begin
        Grille.CellValues[4, Row] := NewValue;
        Compteur := Compteur + 1;
        Modifier := True;
      end;
    end;
    PgiInfo(IntToStr(Compteur) + ' donnée(s) modifiée(s).', Ecran.Caption);
  end
  else
    PgiBox(TexteMessage[3], Ecran.caption);
end;


procedure TOF_PGREAFFECTORGANISME.OnUpdate;
var
  NbMajOk: integer;
  Row: integer;
  St, StWhere: string;
begin
  inherited;
  NbMajOk := 0;
  InitMove(Grille.RowCount, '');
  TEvent := TStringList.create; //PT1
  for Row := 1 to Grille.RowCount - 1 do
    //Recherche par lignes des valeurs modifiées
  begin
    RecupClauseChampReaffecter(3, Row, St);
    if (St <> '') and (Grille.Cellvalues[4, Row] <> '') then
    try
      BeginTrans;
      MoveCur(False);
      StWhere := ' WHERE PCT_NATURERUB="COT" AND ##PCT_PREDEFINI## PCT_RUBRIQUE="' + Grille.Cells[0, Row] + '" ';
      NbMajOk := NbMajOk + ExecuteSql('UPDATE COTISATION SET ' + st + StWhere +
        ' AND PCT_PREDEFINI="' + Grille.CellValues[2, Row] + '"');
      TEvent.Add('Réaffectation de l''organisme de la rubrique ' + Grille.Cells[0, Row]); //PT1
      if UpperCase(Grille.Cells[5, Row]) = 'OUI' then
      begin
        St := ConvertPrefixe(St, 'PCT', 'PHB');
        StWhere := ConvertPrefixe(StWhere, 'PCT', 'PHB');
        StWhere := StWhere + ' AND PHB_DATEFIN>="' + USDateTime(StrToDate(GetControlText('DATEDEB'))) + '" ' +
          'AND PHB_DATEFIN<="' + USDateTime(StrToDate(GetControlText('DATEFIN'))) + '" ';
        MoveCur(False);
        NbMajOk := NbMajOk + ExecuteSql('UPDATE HISTOBULLETIN SET ' + st + ' ' + StWhere);
        // St:=ConvertPrefixe(St,'PHB','PHA'); StWhere:=ConvertPrefixe(StWhere,'PHB','PHA');
        // MoveCur(False);
        // NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOANALPAIE SET '+st+' '+StWhere);
        TEvent.Add('Réaffectation de l''organisme sur l''historique de paie de la rubrique ' + Grille.Cells[0, Row]); //PT1
      end;
      MoveCur(False);
      CommitTrans;
      TEvent.Add('Mise à jour des données OK'); //PT1
      CreeJnalEvt('004', '126', 'OK', nil, nil, TEvent); //PT1
      Modifier := False;
    except
      Rollback;
      NbMajOk := 0;
      TEvent.Add('Une erreur est survenue lors de la mise à jour des données'); //PT1
      CreeJnalEvt('004', '126', 'ERR', nil, nil, TEvent); //PT1
      PGIBox('Une erreur est survenue lors de la mise à jour des données', Ecran.caption);
    end;
  end;
  FiniMove;
  if NbMajOk = 0 then PGIInfo(TexteMessage[8], Ecran.Caption)
  else
  begin
    PGIInfo(IntToStr(NbMajOk) + ' enregistrement(s) sauvegardé(s).Traitement terminé.', Ecran.Caption);
    OnClickBtnActiveWhere(nil);
  end;
  if TEvent <> nil then TEvent.Free;
end;

initialization
  registerclasses([TOF_PGREAFFECT, TOF_PGREAFFECTTABLE, TOF_PGREAFFECTORGANISME]);
end.

