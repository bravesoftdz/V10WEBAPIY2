unit UtilDimArticle;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls
  , HCtrls, HEnt1, EntGc, Hdimension, AGLInit, Windows,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, Fe_Main,
  {$ENDIF}
  {$IFDEF HDIM}
  utilPgi,
  {$ENDIF}
  UTob, UdimArticle, UTofMBOParamDim, UtilArticle;

function OnChangeAfficheChampDimMul(DimensionsArticle: TODimArticle; natureDoc: string): string;
procedure ParamGrille(var DimensionsArticle: TODimArticle; natureDoc, naturePiece: string);
procedure AfficheUserPref(DimensionsArticle: TODimArticle; natureDoc, naturePiece: string);
function RetourneCoordonneeCellule(ACol: integer; GS: THGrid): TPoint; //AC
procedure ChampSetReadOnly(natureDoc: string; numChamp: integer; DimensionsArticle: TODimArticle; var lDim: THDimension);
procedure ControlePresenceChampsObligatoires(var DimensionsArticle: TODimArticle; natureDoc: string);
function SetInitTypeDonnee(NomChamp: string): variant;
function ChampToLength(NomChamp: string): integer;
procedure GetSetStatut(F: TForm; NomChamp: string; GetStatut, bModifCritAuto: boolean; var Valeur: boolean);
//function FindItemToFocused(GridDim: THGrid; Masque: string; TypeMasque: string = '...'): THDimensionItem;
procedure ItemDimSetFocus(DimensionsArticle: TODimArticle; itemDim: THDimensionItem; ArticleDimFocus: string; iChamp: integer = 1; NomChamp: string = '');
procedure FocusedFirstItem(THDim : THDimension);

var TableDim: THDimensionItemList;
  DiOn, DiL1, DiL2, DiC1, DiC2: THValComboBox;
  nbZone: integer;
  zonesImprimees: array[1..11] of string;
  masque: string;
  TitreOngl, TitreLig1, TitreLig2, TitreCol1, TitreCol2: string;
  nomChampCodeArticle: string;

implementation

function OnChangeAfficheChampDimMul(DimensionsArticle: TODimArticle; natureDoc: string): string;
var nbVal, lgMax, ColLength, MaxColLength, iChamp: integer;
  bTitreMajuscule: boolean; // Ajustement largeur de colonnes si majuscules
  Critere, st: string;
begin
  result := '-';
  nbVal := 0;
  lgMax := 0;
  bTitreMajuscule := False;
  st := DimensionsArticle.Dim.TitreValeur;
  while st <> '' do
  begin
    Critere := Trim(ReadTokenSt(st));
    if Critere <> '' then
    begin
      //if TFFiche(ECRAN).Pages.canvas.textWidth (ValMul) > lgMax
      //  then lgMax := TFFiche(ECRAN).Pages.canvas.textWidth (ValMul);
      if lgMax < length(Critere) then
      begin
        lgMax := Length(Critere);
        bTitreMajuscule := (Critere = UpperCase(Critere));
      end;
      inc(nbVal);
    end;
  end;
  if bTitreMajuscule then lgMax := lgMax * 8 else lgMax := lgMax * 5;

  DimensionsArticle.ChangeChampDimMul(True);

  // Ajustement largeur des colonnes en fonction des titres
  with DimensionsArticle.Dim do
  begin
    if nbVal > 10 then nbVal := 10 else if nbVal = 0 then
    begin
      nbVal := 1;
      lgMax := DefaultCol1Width - 10
    end;
    NbValeurs := nbval;
    if nbval > 0 then DefaultColValWidth := lgMax + 10
    else DefaultColValWidth := 64;
    // Agandissement taille de la 1ère colonne si insuffisante
    if nbval = 1 then DefaultCol1Width := DefaultColValWidth;
  end;

  // Ajustement largeur des colonnes de données en fonction du type des champs affichés
  MaxColLength := 0;
  for iChamp := 1 to MaxDimChamp do if DimensionsArticle.NewDimChamp[iChamp] <> '' then
    begin
      ColLength := ChampToLength(DimensionsArticle.NewDimChamp[iChamp]);
      if ColLength > MaxColLength then MaxColLength := ColLength;
    end;
  if MaxColLength <> 0 then MaxColLength := (MaxColLength * 6) else MaxColLength := 50;
  DimensionsArticle.Dim.DefaultColWidth := MaxColLength;

  DimensionsArticle.RefreshGrid(DimensionsArticle.HideUnUsed);
end;

// Gestion des champs en ReadOnly

procedure ChampSetReadOnly(natureDoc: string; numChamp: integer; DimensionsArticle: TODimArticle; var lDim: THDimension);
var NomChamp: string;
begin
  if DimensionsArticle.TableDim = nil then exit;
  NomChamp := DimensionsArticle.NewDimChamp[numChamp];
  if (natureDoc <> NAT_ARTICLE) and ((Copy(NomChamp, 1, 3) = 'GA_') or (NomChamp = 'GL_DP%')) then lDim.ReadOnly[numChamp] := True;
  if (natureDoc = NAT_STOCK) or (natureDoc = NAT_STOVTE) or (natureDoc = NAT_TOBVIEW) then lDim.ReadOnly[numChamp] := True;
  if (Copy(NomChamp, 1, 3) = 'GQ_') then lDim.ReadOnly[numChamp] := True;
end;

// Gestion des champs obligatoires

procedure ControlePresenceChampsObligatoires(var DimensionsArticle: TODimArticle; natureDoc: string);
var QQ: TQuery;
  iChamp, iTab: integer;
  stTemp: string;
begin
  // Saisie des propositions de transferts : gestion de la consultation du dépôt.
  QQ := OpenSQL('select CO_ABREGE,CO_LIBELLE from COMMUN where CO_TYPE="GDO" and CO_LIBRE like "%' + natureDoc + '%" order by CO_CODE desc', True);
  while not QQ.EOF do
  begin
    if QQ.Findfield('CO_ABREGE').AsString <> '' then
    begin
      iChamp := 1;
      stTemp := QQ.Findfield('CO_ABREGE').AsString;
      while (stTemp <> DimensionsArticle.NewDimChamp[iChamp]) and (iChamp < MaxDimChamp) do inc(iChamp);
      if (stTemp <> DimensionsArticle.NewDimChamp[iChamp]) then
      begin
        // Insertion du nom de champ en début de tableau
        for iTab := MaxDimChamp - 1 downto 2
          do DimensionsArticle.NewDimChamp[iTab] := DimensionsArticle.NewDimChamp[iTab - 1];
        DimensionsArticle.NewDimChamp[1] := stTemp;
        // Insertion du libellé du nom de champ en début de tokenString
        stTemp := DimensionsArticle.Dim.TitreValeur;
        DimensionsArticle.Dim.TitreValeur := QQ.Findfield('CO_LIBELLE').AsString + ';' + stTemp;
      end;
    end;
    QQ.next;
  end;
  ferme(QQ);
end;

procedure ParamGrille(var DimensionsArticle: TODimArticle; natureDoc, naturePiece: string);
var iChamp, x, iValeur, iFor: integer;
  stInit, st, stTmp, Critere, ChampMul, ValMul, detail, sauveParamGrille, OldTypeMasque: string;
begin
  OldTypeMasque := DimensionsArticle.TypeMasqueDim;
  if DimensionsArticle.HideUnUsed then detail := '-' else detail := 'X';
  if DimensionsArticle.SauveParamGrille then sauveParamGrille := 'X' else sauveParamGrille := '-';
  {$IFDEF HDIM}
  stInit := 'NATDOC=' + natureDoc + ';NATPIE=' + naturePiece + ';SAUVER=' + sauveParamGrille +
    ';DETAIL=' + detail + ';TYPMAS=' + DimensionsArticle.TypeMasqueDim +
    ';DIMDEP=' + DimensionsArticle.DimDepot + ';';
  {$ELSE}
  stInit := natureDoc + ';' + naturePiece + ';' + sauveParamGrille + ';' + detail + ';' +
    DimensionsArticle.TypeMasqueDim + ';' + DimensionsArticle.DimDepot + ';';
  {$ENDIF}
  for iChamp := 1 to MaxDimChamp do
  begin
    {$IFDEF HDIM}
    stInit := stInit + 'CHAMP' + IntToStr(iChamp) + '=' + DimensionsArticle.NewDimChamp[iChamp] + ';';
    {$ELSE}
    stInit := stInit + DimensionsArticle.NewDimChamp[iChamp] + ';';
    {$ENDIF}
  end;
  {$IFDEF HDIM}
  stInit := stInit + 'MASQUE=' + DimensionsArticle.MasqueDim;
  stInit := stInit + ';AFFDIMTOTFIN=';
  for iFor := 1 to MaxDimension do stInit := stInit + CheckToString(DimensionsArticle.AffDimTotFin[iFor]);
  stInit := stInit + ';AFFDIMDET=';
  for iFor := 1 to MaxDimension do stInit := stInit + CheckToString(DimensionsArticle.AffDimDet[iFor]);
  {$ENDIF}
  DimensionsArticle.ChangeChampDimMul(False); // Svgde valeurs saisies dans TOB

  if (DimensionsArticle.NatureDoc = NAT_PROPSAI) and (not DimensionsArticle.bMasqueDepot) then
  begin
    // Bascule des elts de TOBArticleDim dans TobDimensions
    if DimensionsArticle.TOBArticleDim <> nil then
      while DimensionsArticle.TOBArticleDim.detail.count > 0
        do DimensionsArticle.TOBArticleDim.detail[0].ChangeParent(DimensionsArticle.TobDimensions, -1);
  end;
  if (DimensionsArticle.NatureDoc = NAT_ARTICLE) or (DimensionsArticle.NatureDoc = NAT_STOCK)
    then TheTob := DimensionsArticle.TobPrefEtab;
  if (DimensionsArticle.NatureDoc = NAT_PROPSAI) or (DimensionsArticle.NatureDoc = NAT_PROPDEP)
    then DimensionsArticle.TOBArticleDim := nil;
  st := AglLanceFiche('MBO', 'ARTDIM_PARAM', '', '', stInit);
  if ((DimensionsArticle.NatureDoc = NAT_ARTICLE) or (DimensionsArticle.NatureDoc = NAT_STOCK))
    and (TheTob <> nil) then DimensionsArticle.TobPrefEtab := TheTob;
  iValeur := 0;
  if st <> '' then
  begin
    DimensionsArticle.Dim.TitreValeur := '';
    stTmp := st;
    while stTmp <> '' do
    begin
      Critere := Trim(ReadTokenSt(stTmp));
      if Critere <> '' then
      begin
        x := pos('=', Critere);
        if x <> 0 then ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'DETAIL' then DimensionsArticle.HideUnUsed := (ValMul = '-')
        else if ChampMul = 'SAUVER' then DimensionsArticle.SauveParamGrille := (ValMul = 'X')
        else if ChampMul = 'TYPMAS' then DimensionsArticle.TypeMasqueDim := ValMul
        else if ChampMul = 'DIMDEP' then DimensionsArticle.DimDepot := ValMul
        else if ChampMul = 'VALEUR' then
        begin
          if iValeur = 0 then for iFor := 1 to MaxDimChamp do DimensionsArticle.NewDimChamp[iFor] := '';
          if ValMul <> '' then
          begin
            inc(iValeur);
            DimensionsArticle.NewDimChamp[iValeur] := ValMul;
          end;
        end
        else if (ChampMul = 'TITRE') and (ValMul <> '(Aucun)')
          then DimensionsArticle.Dim.TitreValeur := DimensionsArticle.Dim.TitreValeur + ValMul + ';'
          {$IFDEF HDIM}
        else if ChampMul = 'AFFDIMTOTFIN' then
        begin
          for iFor := 1 to MaxDimension do DimensionsArticle.AffDimTotFin[iFor] := StringToCheck(copy(ValMul, iFor, 1));
        end
        else if ChampMul = 'AFFDIMDET' then
        begin
          for iFor := 1 to MaxDimension do DimensionsArticle.AffDimDet[iFor] := StringToCheck(copy(ValMul, iFor, 1));
        end;
        {$ENDIF}
      end;
    end;
    if iValeur > 0 then
    begin
      ControlePresenceChampsObligatoires(DimensionsArticle, natureDoc);
    end;
  end;
  if ((TheTob <> nil) and (DimensionsArticle.TypeMasqueDim <> VH_GC.BOTypeMasque_Defaut)) or
    (DimensionsArticle.TypeMasqueDim <> OldTypeMasque)
    then DimensionsArticle.LoadDimensions(False);
  if (st <> '') or (TheTob <> nil) or (DimensionsArticle.TypeMasqueDim <> OldTypeMasque) //or (iValeur>0)
  then OnChangeAfficheChampDimMul(DimensionsArticle, natureDoc);
end;

{Affichage des champs par défaut de l'utilisateur}
{Paramètres : entrée = DimensionsArticle
                       natureDoc : fiche article, achat, vente, stock, ..
                       naturePiece : cde client, cde fourn, facture, ..
}

procedure AfficheUserPref(DimensionsArticle: TODimArticle; natureDoc, naturePiece: string);
begin
  OnChangeAfficheChampDimMul(DimensionsArticle, natureDoc);
end;

function RetourneCoordonneeCellule(ACol: Integer; GS: THGrid): TPoint; //AC
var CoordonneCellule: TRect;
begin
  CoordonneCellule := GS.CellRect(ACol, GS.Row);
  Result := GS.ClientToScreen(CoordonneCellule.TopLeft);
  if Result.y <= 372 then CoordonneCellule := GS.CellRect(ACol, GS.Row + 1);
  Result := GS.ClientToScreen(CoordonneCellule.TopLeft);
end;

procedure LibereTableDim;
var iDim: integer;
begin
  if TableDim <> nil then
  begin
    for iDim := 0 to TableDim.count - 1 do
      if TableDim.Items[iDim] <> nil then TableDim.Items[iDim].free;
    TableDim.free;
    TableDim := nil;
  end;
end;

procedure SetTypeDonnee(Dim: THDimension; NomChamp: string; NumChamp: integer);
var Typ: string;
begin
  if NomChamp <> '' then
  begin
    typ := ChampToType(NomChamp);
    if (Typ = 'DOUBLE') or (Typ = 'RATE') then
    begin
      Dim.TypeDonnee[NumChamp] := dotReel;
      Dim.Alignment[NumChamp] := taRightJustify
    end
    else
      if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then
    begin
      Dim.TypeDonnee[NumChamp] := dotReel;
      Dim.Alignment[NumChamp] := taRightJustify
    end
    else
      if Typ = 'DATE' then
    begin
      Dim.TypeDonnee[NumChamp] := dotDate;
      Dim.Alignment[NumChamp] := taRightJustify
    end
    else
    begin
      Dim.TypeDonnee[NumChamp] := dotString;
      Dim.Alignment[NumChamp] := taLeftJustify
    end
  end
  else
  begin
    Dim.TypeDonnee[NumChamp] := dotString;
    Dim.Alignment[NumChamp] := taLeftJustify
  end;
end;

function SetInitTypeDonnee(NomChamp: string): variant;
var Typ: string;
begin
  Result := '';
  if NomChamp <> '' then
  begin
    Typ := ChampToType(NomChamp);
    if (Typ = 'DOUBLE') or (Typ = 'RATE') then Result := 0.0
    else if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then Result := 0
    else if Typ = 'DATE' then Result := '01/01/1900'
    else Result := '';
  end;
end;

function ChampToLength(NomChamp: string): integer;
var Typ: string;
begin
  Result := 0;
  if NomChamp = 'GF_CALCULREMISE' then Result := 12 // Spécif Tarif
  else if NomChamp <> '' then
  begin
    Typ := ChampToType(NomChamp);
    if (copy(Typ, 1, 7) = 'VARCHAR') then Result := StrToInt(copy(Typ, 9, length(Typ) - 9))
    else if (Typ = 'DOUBLE') or (Typ = 'RATE') then Result := 12
    else if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then Result := 8
    else if Typ = 'DATE' then Result := 10
    else Result := 8; // COMBO,BOOLEAN,BLOB,...
    // Limitation à 20 -> varchar(100), ...
    if Result > 20 then Result := 20;
  end;
end;

procedure GetSetStatut(F: TForm; NomChamp: string; GetStatut, bModifCritAuto: boolean; var Valeur: boolean);
var statut: TCheckBox;
begin
  statut := TCheckBox(F.FindComponent(NomChamp));
  if (GetStatut) then // GetStatut=True -> svgde valeur actuelle
  begin
    if not bModifCritAuto then
    begin
      Valeur := statut.checked;
      statut.checked := Boolean(NomChamp = 'STATUT_DIM');
    end;
  end
  else statut.checked := Valeur;
end;

{***********A.G.L.***********************************************
Auteur  ...... : J DITTMAR
Créé le ...... : 16/10/2003
Modifié le ... : 16/10/2003
Description .. : Recherche de l'item devant recevoir le focus = celui le + en 
Suite ........ : haut à gauche
Mots clefs ... : THDIMENSION;ITEMDIM;FOCUS
*****************************************************************}
procedure FocusedFirstItem(THDim : THDimension);
Var st1, st2, st3, st4, st5 : string;
begin
  With THDIM do
  begin
    if DimOngl <> Nil then st1 := DimOngl.Values[0] else St1 := '';
    if DimLig1 <> Nil then st2 := DimLig1.Values[0] else St2 := '';
    if DimLig2 <> Nil then st3 := DimLig2.Values[0] else St3 := '';
    if DimCol1 <> Nil then st4 := DimCol1.Values[0] else St4 := '';
    if DimCol2 <> Nil then st5 := DimCol2.Values[0] else St5 := '';
    // Positione le focus sur la première cellule du premier onglet
    GoToValeurDimension(st1, st2, st3, st4, st5, False, 0, False);
    // Si la valeur n'existe pas
    if (DimG.Objects[DimG.Col, DimG.Row] = nil) or (THDimensionItem(DimG.Objects[DimG.Col, DimG.Row]).ToDelete) then
      // Recherche la première valeur existante
      GotoNextActiveCell(sdHorizontal, False);
    // Passe en mode édition
    UpdateEditing;
    // Donne le focus à l'objet
    GridDim.SetFocus;
  end;
end;

{  Remplacer par la procedure FocusedFirstItem
function FindItemToFocused(GridDim: THGrid; Masque: string; TypeMasque: string = '...'): THDimensionItem;
var iCol, iLig, iPos, iDim: integer;
  ItemFocus, Item: THDimensionItem;
  QQ: TQuery;
  TobDim, TobItem: Tob;
  Found: boolean;
  stPos, Sql: string;
  TypeDim, GrilleDim: array[1..4] of string;
  Rang: array[1..4] of integer;
begin
  Result := nil;
  QQ := OpenSQL('select * from DIMMASQUE where GDM_MASQUE="' + Masque + '" and GDM_TYPEMASQUE="' + TypeMasque + '"', True);
  if QQ.EOF then
  begin
    Ferme(QQ);
    exit;
  end;
  for iDim := 1 to 4 do
  begin
    TypeDim[iDim] := '';
    GrilleDim[iDim] := '';
    Rang[iDim] := 99999
  end;
  for iPos := 1 to 6 do
  begin
    stPos := QQ.Findfield('GDM_POSITION' + IntToStr(iPos)).AsString;
    if stPos = 'LI1' then
    begin
      TypeDim[1] := 'DI' + IntToStr(iPos);
      GrilleDim[1] := QQ.Findfield('GDM_TYPE' + IntToStr(iPos)).AsString;
    end
    else if stPos = 'LI2' then
    begin
      TypeDim[2] := 'DI' + IntToStr(iPos);
      GrilleDim[2] := QQ.Findfield('GDM_TYPE' + IntToStr(iPos)).AsString;
    end
    else if stPos = 'CO1' then
    begin
      TypeDim[3] := 'DI' + IntToStr(iPos);
      GrilleDim[3] := QQ.Findfield('GDM_TYPE' + IntToStr(iPos)).AsString;
    end
    else if stPos = 'CO2' then
    begin
      TypeDim[4] := 'DI' + IntToStr(iPos);
      GrilleDim[4] := QQ.Findfield('GDM_TYPE' + IntToStr(iPos)).AsString;
    end;
  end;
  Ferme(QQ);

  // Chargement en tob des CODEDIM et RANG des grilles de dimension
  Sql := 'select GDI_TYPEDIM,GDI_GRILLEDIM,GDI_CODEDIM,GDI_RANG from DIMENSION where (1=3)';
  for iDim := 1 to 4 do if TypeDim[iDim] <> '' then
      Sql := Sql + ' or (GDI_TYPEDIM="' + TypeDim[iDim] + '" and GDI_GRILLEDIM="' + GrilleDim[iDim] + '")';
  QQ := OpenSQL(Sql, True);
  if QQ.EOF then exit;
  TobDim := TOB.CreateDB('_DIMENSION', nil, -1, QQ);
  TobDim.LoadDetailDB('_DIMENSION', '', '', QQ, False);
  Ferme(QQ);

  ItemFocus := nil;
  for iLig := GridDim.FixedRows to GridDim.FixedRows + GridDim.VisibleRowCount - 1 do
  begin
    for iCol := GridDim.FixedCols to GridDim.FixedCols + GridDim.VisibleColCount - 1 do
    begin
      Item := THDimensionItem(GridDim.Objects[iCol, iLig]);
      if Item <> nil then
      begin
        // Si les rangs de l'item dans les grilles de dimensions sont inférieurs, l'item est retenu.
        iDim := 1;
        Found := False;
        while iDim < 4 do
        begin
          if TypeDim[iDim] <> '' then
          begin
            TobItem := TobDim.FindFirst(['GDI_GRILLEDIM', 'GDI_CODEDIM'], [GrilleDim[iDim], Tob(Item.data).GetValue('GA_CODEDIM' + Copy(TypeDim[iDim], 3, 1))],
              False);
            if TobItem <> nil then
            begin
              // Si item retenu, mémorisation de tous ses rangs !
              if Found then Rang[iDim] := TobItem.GetValue('GDI_RANG')
              else if TobItem.GetValue('GDI_RANG') < Rang[iDim] then
              begin
                Found := True;
                ItemFocus := Item;
                Rang[iDim] := TobItem.GetValue('GDI_RANG');
              end
              else if TobItem.GetValue('GDI_RANG') > Rang[iDim] then break;
            end;
          end;
          inc(iDim);
        end;
      end;
    end;
  end;
  if TobDim <> nil then TobDim.Free;
  Result := ItemFocus;
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Positionne le focus dans le THDimension sur :
Suite ........ : - itemDim s'il est renseigné (<>nil)
Suite ........ : - l'itemDim dont GA_ARTICLE = ArticleDimFocus sinon
Mots clefs ... : FOCUS;ITEMDIM;THDIM
*****************************************************************}
procedure ItemDimSetFocus(DimensionsArticle: TODimArticle; itemDim: THDimensionItem; ArticleDimFocus: string; iChamp: integer = 1; NomChamp: string = '');
var iItem, index: integer;
begin
  if NomChamp <> '' then
  begin
    iChamp := 1;
    while (iChamp < MaxDimChamp) and (DimensionsArticle.NewDimChamp[iChamp] <> NomChamp) do inc(iChamp);
    if DimensionsArticle.NewDimChamp[iChamp] <> NomChamp then iChamp := 1;
  end;
  if itemDim = nil then
  begin
    iItem := 0; index := -1;
    while (index = -1) and (iItem < DimensionsArticle.TableDim.count) do
    begin
      itemDim := THDimensionItem(DimensionsArticle.TableDim.Items[iItem]);
      if (TOB(ItemDim.data) <> nil) and (TOB(ItemDim.data).GetValue('GA_ARTICLE') = ArticleDimFocus)
        then index := iItem;
      inc(iItem);
    end;
    if index < 0 then index := 0;
    itemDim := DimensionsArticle.TableDim.Items[index];
  end;
  DimensionsArticle.Dim.FocusDim(ItemDim, iChamp);
  DimensionsArticle.Dim.UpdateEditing;
  DimensionsArticle.Dim.GridDim.SetFocus;
end;


end.
