{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/06/2003
Modifié le ... : 07/08/2003
Description .. : Cubes des achats et des ventes
Mots clefs ... : CUBE;STAT;DIMENSION
*****************************************************************}
unit UTofMBOVentesCube;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTOF,
  StdCtrls, HCtrls, HDimension, AglInitGC, M3FP, hmsgbox,
  {$IFDEF EAGLCLIENT}
  emul, utileAGL, Maineagl,
  {$ELSE}
  mul, dbTables, FE_Main,
  {$ENDIF}
  EntGC, HEnt1, utilPGI, UtilGC, UtilArticle, Cube;

type
  TOF_MBOVentesCube = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  private
    FLoaded : boolean;
    procedure MyInitFormule;
    procedure AfficheDim;
  end;

implementation

procedure TOF_MBOVentesCube.AfficheDim;
var iCol: integer;
  Title, Name, Rang, st: string;
  cb: TCheckBox;
  F: TFCube;
begin
  F := TFCube (Ecran);
  st := '';
  for iCol := 1 to MaxDimension do
  begin
    st := 'GA_GRILLEDIM' + IntToStr(iCol);
    if TGroupBox(GetControl('GBDIM')).Visible = True then
    begin
      cb := TCheckBox(GetControl('CB' + InttoStr(iCol)));
      Title := cb.Caption;
      if Title='.-' then continue; //EPZ
      Name := 'YYY_' + StrToAlias(Uppercase(Title)); // préfixage YYY_ obligatoire
      Rang := 'YYY_TRI_' + StrToAlias(Uppercase(Title));
      // si la dimension existe déjà mais qu'elle n'est pas cochée alors on la supprime
      if F.ExistFormule(Name) and not cb.checked then
      begin
        F.DelFormule(Name);
        F.DelFormule(Rang);
      end else
        // si la dimension n'existe pas mais qu'elle est cochée alors on la rajoute
        if not F.ExistFormule(Name) and cb.checked then
      begin
        F.AddFormule(Title, 'GDI' + IntToStr(iCol) + '.GDI_LIBELLE', Name);
        F.AddFormule('Rang Dim' + IntToStr(iCol), 'GDI' + IntToStr(iCol) + '.GDI_RANG', Rang);
      end;
      // suppression des champs grille de dimension pour les dimensions non utilisées
      if F.ExistFormule(st) and not cb.checked then
        F.DelFormule(st) else
        if not F.ExistFormule(st) and cb.checked then
        F.AddFormule(ChampToLibelle(st), '', st);
    end else
      F.DelFormule(st); // s'il y a un filtre avec des dimensions on ne les fait pas apparaître
  end;
end;

procedure TOF_MBOVentesCube.MyInitFormule;
var iCol: integer;
  Title, Name, Rang: string;
  cb: TCheckBox;
  F: TFCube;
begin
  if FLoaded then exit;
  FLoaded := TRUE;
  F := TFCube(Ecran);
  for iCol := 1 to MaxDimension do
    if TGroupBox(GetControl('GBDIM')).Visible = True then
    begin
      cb := TCheckBox(GetControl('CB' + InttoStr(iCol)));
      Title := cb.Caption;
      if Title='.-' then continue; //EPZ
      Name := 'YYY_' + StrToAlias(Uppercase(Title)); // préfixage YYY_ obligatoire
      Rang := 'YYY_TRI_' + StrToAlias(Uppercase(Title));
      F.AddFormule(Title, 'GDI' + IntToStr(iCol) + '.GDI_LIBELLE', Name);
      F.AddFormule('Rang Dim' + IntToStr(iCol), 'GDI' + IntToStr(iCol) + '.GDI_RANG', Rang);
    end;
end;

procedure TOF_MBOVentesCube.OnArgument(Arguments: string);
var F: TFCube;
  TCB: TCheckBox;
  iCol, x: integer;
  Critere, NomCritere, ValCritere, LibType, LibDim: string;
  {$IFDEF NOMADE}
  stPlus, stNatDef: string;
  {$ENDIF}
begin
  inherited;
  F := TFCube(Ecran);
  FLoaded := FALSE;
  {récup des arguments }
  repeat
    Critere := Trim(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        NomCritere := uppercase(copy(Critere, 1, x - 1));
        ValCritere := copy(Critere, x + 1, length(Critere));
        if NomCritere = 'TYPECUB' then LibType := ValCritere
        else
          if NomCritere = 'TYPEART' then LibDim := ValCritere;
      end;
    end;
  until Critere = '';
  if (ctxMode in V_PGI.PGIContexte) then
  begin
    SetControlProperty('GL_DEPOT', 'Plus', 'GDE_SURSITE="X"');
    SetControlProperty('GL_ETABLISSEMENT', 'Plus', 'ET_SURSITE="X"');
  end;
  // Paramétrage des libellés des familles, stat. article
  ChangeLibre2('TGA_COLLECTION', Ecran);
  for iCol := 1 to 3 do ChangeLibre2('TGA_FAMILLENIV' + InttoStr(iCol), Ecran);
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
  begin
    for iCol := 4 to 8 do ChangeLibre2('TGA2_FAMILLENIV' + InttoStr(iCol), Ecran);
    for iCol := 1 to 2 do ChangeLibre2('TGA2_STATART' + InttoStr(iCol), Ecran);
  end
  else
  begin
    for iCol := 4 to 8 do
    begin
      SetControlVisible('TGA2_FAMILLENIV' + InttoStr(iCol), False);
      SetControlVisible('GA2_FAMILLENIV' + InttoStr(iCol), False);
    end;
    for iCol := 1 to 2 do
    begin
      SetControlVisible('TGA2_STATART' + InttoStr(iCol), False);
      SetControlVisible('GA2_STATART' + InttoStr(iCol), False);
    end;
  end;

  if LibDim = 'DIM' then
  begin
    // Ajout des dimensions sélectionnées dans les champs disponibles
    F.OnInitFormules := MyInitFormule;
    SetControlVisible('GBDIM', true);
    if F.Name = 'ACHATCUB' then
      F.Caption := TraduireMemoire('Statistiques croisées des achats à la dimension') else
      F.Caption := TraduireMemoire('Statistiques croisées des ventes à la dimension');
    for iCol := 1 to MaxDimension do
    begin
      TCB := TCheckBox(GetControl('CB' + InttoStr(iCol)));
      TCB.Caption := RechDom('GCCATEGORIEDIM', 'DI' + InttoStr(iCol), False);
      if TCB.Caption = '.-' then TCB.Visible := False;
    end;
  end else
  begin
    if F.Name = 'ACHATCUB' then
      F.Caption := TraduireMemoire('Statistiques croisées des achats par article') else
      F.Caption := TraduireMemoire('Statistiques croisées des ventes par article');
  end;
  UpdateCaption(F);

  {$IFDEF NOMADE}
  stPlus := ' AND GPP_VENTEACHAT="' + LibType + '"';
  if LibType = 'VEN' then
  begin
    stPlus := stPlus + ' AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")';
    stNatDef := 'CC';
  end else
  begin
    stPlus := stPlus + ' AND (GPP_NATUREPIECEG="CF" OR GPP_TRSFACHAT="X")';
    stNatDef := 'CF';
  end;
  SetControlProperty('GL_NATUREPIECEG', 'PLUS', stPlus);
  SetControlText('GL_NATUREPIECEG', stNatDef);
  {$ELSE}
  if LibType = 'VEN' then
  begin
    if (VH_GC.GCSeria = True) or (V_PGI.VersionDemo = True)
      then SetControlText('GL_NATUREPIECEG', 'FAC;AVC;FFO')
    else SetControlText('GL_NATUREPIECEG', 'FFO');
  end else
  begin
    SetControlText('GL_NATUREPIECEG', 'BLF');
  end;
  {$ENDIF}

  // Paramétrage des libellés des tables libres
  GCMAJChampLibre(F, False, 'COMBO', 'GA_LIBREART', 10, '');
  GCMAJChampLibre(F, False, 'COMBO', 'ET_LIBREET', 10, '');
  GCMAJChampLibre(F, False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '');
  if VH_GC.GCMultiDepots then GCMAJChampLibre(F, False, 'COMBO', 'GDE_LIBREDEP', 10, '')
  else SetControlVisible('GB_TABLELIBRDEP', False);

  // Mise en forme des libellés des dates, booléans libres et montants libres des articles
  if (GCMAJChampLibre(F, False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False);
  if (GCMAJChampLibre(F, False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False);
  if (GCMAJChampLibre(F, False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False);

  // Mise en forme des libellés des dates, booléans libres et montants libres des établissements
  if (GCMAJChampLibre(F, False, 'EDIT', 'ET_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALETAB', False);
  if (GCMAJChampLibre(F, False, 'EDIT', 'ET_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEETAB', False);
  if (GCMAJChampLibre(F, False, 'BOOL', 'ET_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLETAB', False);

  // Mise en forme des libellés des dates, booléans libres, montants libres et textes libres des clients
  if (GCMAJChampLibre(F, False, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL1', False);
  if (GCMAJChampLibre(F, False, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE1', False);
  if (GCMAJChampLibre(F, False, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL1', False);
  if (GCMAJChampLibre(F, False, 'EDIT', 'YTC_TEXTELIBRE', 3, '') = 0) then SetControlVisible('GB_TEX1', False);

  if VH_GC.GCMultiDepots then
  begin
    // Mise en forme des libellés des dates, booléans libres et montants libres des dépôts
    if (GCMAJChampLibre(F, False, 'EDIT', 'GDE_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALDEP', False);
    if (GCMAJChampLibre(F, False, 'EDIT', 'GDE_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEDEP', False);
    if (GCMAJChampLibre(F, False, 'BOOL', 'GDE_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLDEP', False);
  end else
  begin
    SetControlVisible('GB_VALDEP', False);
    SetControlVisible('GB_DATEDEP', False);
    SetControlVisible('GB_BOOLDEP', False);
  end;
end;

procedure TOF_MBOVentesCube.OnLoad;
var stDim, xx_where: string;
  iCol: integer;
  F: TFCube;
begin
  F := TFCube(Ecran);
  AfficheDim;
  // Récupération des zones libres
  xx_where := '';
  xx_where := GCXXWhereChampLibre(F, xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'MULTICOMBO', 'GA_FAMILLENIV', 3, '');
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
  begin
    xx_where := GCXXWhereChampLibre(F, xx_where, 'MULTICOMBO', 'GA2_FAMILLENIV', 5, '');
    xx_where := GCXXWhereChampLibre(F, xx_where, 'MULTICOMBO', 'GA2_STATART', 2, '');
  end;
  xx_where := GCXXWhereChampLibre(F, xx_where, 'MULTICOMBO', 'GA_LIBREART', 10, '');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'MULTICOMBO', 'ET_LIBREET', 10, '');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'BOOL', 'ET_BOOLLIBRE', 3, '');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'DATE', 'ET_DATELIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'EDIT', 'ET_VALLIBRE', 3, '_');

  if VH_GC.GCMultiDepots then
  begin
    xx_where := GCXXWhereChampLibre(F, xx_where, 'MULTICOMBO', 'GDE_LIBREDEP', 10, '');
    xx_where := GCXXWhereChampLibre(F, xx_where, 'BOOL', 'GDE_BOOLLIBRE', 3, '');
    xx_where := GCXXWhereChampLibre(F, xx_where, 'DATE', 'GDE_DATELIBRE', 3, '_');
    xx_where := GCXXWhereChampLibre(F, xx_where, 'EDIT', 'GDE_VALLIBRE', 3, '_');
  end;

  xx_where := GCXXWhereChampLibre(F, xx_where, 'BOOL', 'YTC_BOOLLIBRE', 3, '');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'DATE', 'YTC_DATELIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'EDIT', 'YTC_VALLIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'STRING', 'YTC_TEXTELIBRE', 3, '');

  SetControlText('XX_WHERE', xx_where);

  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
    F.FromSQL := 'LIGNE left join ARTICLE on GA_ARTICLE=GL_ARTICLE'
      + ' left join ARTICLECOMPL on GA2_ARTICLE=GL_ARTICLE'
      + ' left join TIERSCOMPL on YTC_TIERS=GL_TIERS'
      + ' left join DEPOTS on GDE_DEPOT=GL_DEPOT'
      + ' left join ETABLISS on ET_ETABLISSEMENT=GL_ETABLISSEMENT'
  else
    F.FromSQL := 'LIGNE left join ARTICLE on GA_ARTICLE=GL_ARTICLE'
      + ' left join TIERSCOMPL on YTC_TIERS=GL_TIERS'
      + ' left join DEPOTS on GDE_DEPOT=GL_DEPOT'
      + ' left join ETABLISS on ET_ETABLISSEMENT=GL_ETABLISSEMENT';
  if TGroupBox(GetControl('GBDIM')).Visible then
    for iCol := 1 to MaxDimension do
    begin
      stDim := IntToStr(iCol);
      if TCheckBox(GetControl('CB' + IntToStr(iCol))).Checked then
        F.FromSQL := F.FromSQL +
          ' left join DIMENSION GDI' + stDim + ' on  GDI' + stDim + '.GDI_TYPEDIM="DI' + stDim + '" AND' +
          ' GA_GRILLEDIM' + stDim + '=GDI' + stDim + '.GDI_GRILLEDIM AND GA_CODEDIM' + stDim + '=GDI' + stDim + '.GDI_CODEDIM';
    end;

end;

initialization
  RegisterClasses([TOF_MBOVentesCube]);
end.
