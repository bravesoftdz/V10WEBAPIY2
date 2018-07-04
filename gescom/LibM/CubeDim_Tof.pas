{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 21/05/2002
Modifié le ... : 05/06/2003
Description .. : Cube des stocks
Suite ........ :  - par article
Suite ........ :  - par dimension
Mots clefs ... : BOS5;CUBE;DIMENSION;
*****************************************************************}
unit CubeDim_Tof;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTOF,
{$IFDEF EAGLCLIENT}
  eMUL, MaineAgl, 
{$ELSE}
  MUL, dbTables, FE_Main,
{$ENDIF}
  StdCtrls, HCtrls, HDimension, AglInitGC, M3FP, hmsgbox,
  EntGC, HEnt1, UtilPGI, UtilGC, UtilArticle, Cube, ExtCtrls;

type
  TOF_CubeDim = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  private
    bouletun, Dim: boolean;
    procedure MyInitFormule;
  end;

implementation

procedure TOF_CubeDim.MyInitFormule;
var iCol: integer;
  Title, Name, Rang, st: string;
  cb: TCheckBox;
  F: TFCube;
begin
  if bouletun then exit;
  bouletun := true;
  F := TFCube(Ecran);
  st := '';
  for iCol := 1 to MaxDimension do
  begin
    st := 'GA_GRILLEDIM' + IntToStr(iCol);
    if TGroupBox(GetControl('GBDIM')).Visible = True then
    begin
      cb := TCheckBox(GetControl('CB' + InttoStr(iCol)));
      Title := cb.Caption;
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

procedure TOF_CubeDim.OnArgument(Arguments: string);
var iCol: integer;
  TCB: TCheckBox;
  F: TFCube;
begin
  inherited;
  bouletun := true;
  Dim := (Arguments = 'DIM');
  F := TFCube(Ecran);
  if (ctxMode in V_PGI.PGIContexte) then
    SetControlProperty('GQ_DEPOT', 'Plus', 'GDE_SURSITE="X"');

  // Paramétrage des libellés des familles, stat. article
  ChangeLibre2('TGA_COLLECTION', Ecran);
  for iCol := 1 to 3 do ChangeLibre2('TGA_FAMILLENIV' + InttoStr(iCol), Ecran);
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
  begin
    for iCol := 4 to 8 do ChangeLibre2('TGA2_FAMILLENIV' + InttoStr(iCol), Ecran);
    for iCol := 1 to 2 do ChangeLibre2('TGA2_STATART' + InttoStr(iCol), Ecran);
  end else
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

  {if not VH_GC.GCMultiDepots then
  begin
    // Mise en forme des libellés des différentes tables et zones libres quand on est en mono-dépôt
    TGroupBox (TFCube(F).FindComponent ('GB_TABLELIBRDEP')).Caption := 'Tables libres établissement' ;
    TGroupBox (TFCube(F).FindComponent ('GB_VALDEP')).Caption := 'Valeurs établissement' ;
    TGroupBox (TFCube(F).FindComponent ('GB_DATEDEP')).Caption := 'Dates établissement' ;
    TGroupBox (TFCube(F).FindComponent ('GB_BOOLDEP')).Caption := 'Décisions établissement' ;
  end; }

  // Si on travaille à la dimension : mise en forme des libellés des dimensions
  if Dim then
  begin
    // Ajout des dimensions sélectionnées dans les champs disponibles
    F.OnInitFormules := MyInitFormule;
    SetControlVisible('GBDIM', true);
    F.Caption := TraduireMemoire('Statistiques croisées des stocks à la dimension');
    for iCol := 1 to MaxDimension do
    begin
      TCB := TCheckBox(GetControl('CB' + InttoStr(iCol)));
      TCB.Caption := RechDom('GCCATEGORIEDIM', 'DI' + InttoStr(iCol), False);
      if TCB.Caption = '.-' then TCB.Visible := False;
    end;
  end else F.Caption := TraduireMemoire('Statistiques croisées des stocks par article');
  UpdateCaption(F);
  ;

  // Paramétrage des libellés des tables libres articles et dépôts
  GCMAJChampLibre(F, False, 'COMBO', 'GA_LIBREART', 10, '');
  GCMAJChampLibre(F, False, 'COMBO', 'ET_LIBREET', 10, '');
  if VH_GC.GCMultiDepots then GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GDE_LIBREDEP', 10, '')
  else SetControlVisible('GB_TABLELIBRDEP', False);

  // Mise en forme des libellés des dates, booléans libres et montants libres des articles
  if (GCMAJChampLibre(F, False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False);
  if (GCMAJChampLibre(F, False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False);
  if (GCMAJChampLibre(F, False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False);

  // Mise en forme des libellés des dates, booléans libres et montants libres des dépôts
  if (GCMAJChampLibre(F, False, 'EDIT', 'ET_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALETAB', False);
  if (GCMAJChampLibre(F, False, 'EDIT', 'ET_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEETAB', False);
  if (GCMAJChampLibre(F, False, 'BOOL', 'ET_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLETAB', False);
  if VH_GC.GCMultiDepots then
  begin
    // Mise en forme des libellés des dates, booléans libres et montants libres des dépôts
    if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GDE_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALDEP', False);
    if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GDE_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEDEP', False);
    if (GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'GDE_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLDEP', False);
  end
  else
  begin
    SetControlVisible('GB_VALDEP', False);
    SetControlVisible('GB_DATEDEP', False);
    SetControlVisible('GB_BOOLDEP', False);
  end;
end;

procedure TOF_CubeDim.OnLoad;
var xx_where, stDim: string;
  iCol: integer;
  F: TFCube;
begin
  F := TFCube(Ecran);
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
  SetControlText('XX_WHERE', xx_where);

  bouletun := false;
  MyInitFormule;
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
    F.FromSQL := 'DISPO left join ARTICLE on GA_ARTICLE=GQ_ARTICLE left join ARTICLECOMPL on GA2_ARTICLE=GQ_ARTICLE left join DEPOTS on GDE_DEPOT=GQ_DEPOT'
  else
    F.FromSQL := 'DISPO left join ARTICLE on GA_ARTICLE=GQ_ARTICLE left join DEPOTS on GDE_DEPOT=GQ_DEPOT';
  if Dim then
    for iCol := 1 to MaxDimension do
    begin
      stDim := IntToStr(iCol);
      if TCheckBox(GetControl('CB' + IntToStr(iCol))).Checked then
        F.FromSQL := F.FromSQL +
          ' left join DIMENSION GDI' + stDim + ' on  GDI' + stDim + '.GDI_TYPEDIM="DI' + stDim + '" AND' +
          ' GA_GRILLEDIM' + stDim + '=GDI' + stDim + '.GDI_GRILLEDIM AND GA_CODEDIM' + stDim +
          '=GDI' + stDim + '.GDI_CODEDIM';
    end;
end;

initialization
  RegisterClasses([TOF_CubeDim]);
end.
