{***********UNITE*************************************************
Auteur  ...... : Olvier TARCY
Créé le ...... : 21/07/2003
Modifié le ... : 06/08/2003
Description .. : Source TOF de la FICHE : CONSOLIDATION (MBO)
Suite ........ : Etat de consolidation des échanges
Mots clefs ... : TOF;CONSOLIDATION
*****************************************************************}
unit CONSOLIDATION_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  dialogs,
  UTOF,
  ParamSoc,
  HDimension,
  EntGC,
  Math,
  Graphics,
  Utob,
  Vierge,
  M3FP,
  HMsgBox;
//  voirtob;

type
  TOF_CONSOLIDATION = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure OnClick(Sender: TObject);
    procedure AnalyseRepertoire(Rep, Mask: string; TOB_temp: tob);
    function CreateTob: boolean;
    procedure InitLesTob(iMax: integer);
    function SiteActif(CodeSite: string): boolean;
    procedure ChargeCombos;
    procedure SQL2Combo(const NomTable, ChampCode, ChampLibelle, Where, OrderBy: string; Items, Valeurs: HTStrings);
    procedure DefineFont;
    procedure Affichage(TypeAff: boolean = true);
  private
    CD, CT, CRC, CRB, CECH, CARCH, CREJET: TCheckBox;
    REP_DEPART, REP_TRAITE, REP_REJET, REP_ECH, REP_ARCH, REP_REJ: string;
    TOB_MERE, TOB_DEPART, TOB_TRAITE, TOB_REJET_SC, TOB_REJET_BTQ, TOB_JOURS: tob;
    DIM_SITES, DIM_BIN: THValComboBox;
    CONSO: THDimension;
    MasquerSitesInactifs, MasquerSitesOK, DisplayMode: boolean; // DisplayMode: True = Sites -> Corbeiles False = Corbeillles -> Sites
    LesItems, LesValeurs: HTStrings;
    ErrorFont, GoodFont: TFont;
  end;

implementation

procedure TOF_CONSOLIDATION.OnNew;
begin
  inherited;
end;

procedure TOF_CONSOLIDATION.OnDelete;
begin
  inherited;
end;

procedure TOF_CONSOLIDATION.OnUpdate;
var Parent: TForm;
begin
  inherited;
  //  if CONSO=nil then
  //  begin
  CONSO := THDimension(GetControl('TABLEAU'));
  Parent := TForm(CONSO.OWner);
  DIM_BIN := THValComboBox.Create(Parent);
  DIM_BIN.Parent := Parent;
  DIM_BIN.Visible := False;

  DIM_SITES := THValComboBox.Create(Parent);
  DIM_SITES.Parent := Parent;
  DIM_SITES.Visible := False;
  if CD.Checked then
  begin
    DIM_BIN.items.Add(TraduireMemoire('A envoyer'));
    DIM_BIN.values.Add('1');
  end;
  if CT.Checked then
  begin
    DIM_BIN.items.Add(TraduireMemoire('Reçus'));
    DIM_BIN.values.Add('2');
  end;
  if CRC.Checked then
  begin
    DIM_BIN.items.Add(TraduireMemoire('Rejetés au central'));
    DIM_BIN.values.Add('3');
  end;
  if CRB.Checked then
  begin
    DIM_BIN.items.Add(TraduireMemoire('Rejetés en boutique'));
    DIM_BIN.values.Add('4');
  end;
  DIM_BIN.items.Add(TraduireMemoire('Nb jours en erreur'));
  DIM_BIN.values.Add('5');
  //  end;
  Affichage(TCheckBox(GetControl('DISPLAYMODE')).Checked);
end;

procedure TOF_CONSOLIDATION.OnLoad;
begin
  inherited;
end;

procedure TOF_CONSOLIDATION.OnArgument(S: string);
begin
  inherited;
  { Récupération des chemins des corbeilles }
  CD := TCheckBox(GetControl('CB_CD'));
  CD.OnClick := OnClick;
  CT := TCheckBox(GetControl('CB_CT'));
  CT.OnClick := OnClick;
  CRC := TCheckBox(GetControl('CB_CR'));
  CRC.OnClick := OnClick;
  CRB := TCheckBox(GetControl('CB_BTQ'));
  CRB.OnClick := OnClick;
  CECH := TCheckBox(GetControl('CB_ECH'));
  CECH.OnClick := OnClick;
  CARCH := TCheckBox(GetControl('CB_ARCH'));
  CARCH.OnClick := OnClick;
  CREJET := TCheckBox(GetControl('CB_REJET'));
  CREJET.OnClick := OnClick;
  (*
  if TX_LesSites.LeSiteLocal <> nil then
  begin
    REP_DEPART := TX_LesSites.LeSiteLocal.SSI_CDEPART;
    REP_TRAITE := TX_LesSites.LeSiteLocal.SSI_CTRAITE;
    REP_REJET := TX_LesSites.LeSiteLocal.SSI_CREJET;
  end;
  *)
  REP_ECH := GetParamSoc('SO_GCREPORLI');
  REP_ARCH := GetParamSoc('SO_GCREPORLIARCH');
  REP_REJ := GetParamSoc('SO_GCREPORLIREJET');
  if CD.Checked then
    SetControlText('LABEL_CD', REP_DEPART) else
    SetControlText('LABEL_CD', '');
  if CT.Checked then
    SetControlText('LABEL_CT', REP_TRAITE) else
    SetControlText('LABEL_CT', '');
  if CRC.Checked then
    SetControlText('LABEL_CR', REP_REJET) else
    SetControlText('LABEL_CR', '');
  if CECH.Checked then
    SetControlText('LABEL_ECH', REP_ECH) else
    SetControlText('LABEL_ECH', '');
  if CARCH.Checked then
    SetControlText('LABEL_ARCH', REP_ARCH) else
    SetControlText('LABEL_ARCH', '');
  if CREJET.Checked then
    SetControlText('LABEL_REJET', REP_REJ) else
    SetControlText('LABEL_REJET', '');
end;

procedure TOF_CONSOLIDATION.OnClose;
begin
  inherited;
  FreeAndNil(TOB_DEPART);
  FreeAndNil(TOB_TRAITE);
  FreeAndNil(TOB_REJET_SC);
  FreeAndNil(TOB_REJET_BTQ);
  FreeAndNil(TOB_JOURS);
  FreeAndNil(TOB_MERE);
  FreeAndNil(DIM_BIN);
  FreeAndNil(DIM_SITES);
  FreeAndNil(ErrorFont);
  FreeAndNil(GoodFont);
  FreeAndNil(LesItems);
  FreeAndNil(LesValeurs);
end;

procedure TOF_CONSOLIDATION.OnDisplay();
begin
  inherited;
end;

procedure TOF_CONSOLIDATION.OnCancel();
begin
  inherited;
end;

procedure TOF_CONSOLIDATION.OnClick(Sender: TObject);
begin
  if TRadioButton(GetControl('RBT_PGI')).Checked then
  begin
    if CD.Checked then
      SetControlText('LABEL_CD', REP_DEPART) else
      SetControlText('LABEL_CD', '');
    if CT.Checked then
      SetControlText('LABEL_CT', REP_TRAITE) else
      SetControlText('LABEL_CT', '');
    if CRC.Checked then
      SetControlText('LABEL_CR', REP_REJET) else
      SetControlText('LABEL_CR', '');
  end else
    if TRadioButton(GetControl('RBT_GPAO')).Checked then
  begin
    if CECH.Checked then
      SetControlText('LABEL_ECH', REP_ECH) else
      SetControlText('LABEL_ECH', '');
    if CARCH.Checked then
      SetControlText('LABEL_ARCH', REP_ARCH) else
      SetControlText('LABEL_ARCH', '');
    if CREJET.Checked then
      SetControlText('LABEL_REJET', REP_REJ) else
      SetControlText('LABEL_REJET', '');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 21/07/2003
Modifié le ... : 30/07/2003
Description .. : Parcours d'un répertoire
Mots clefs ... : CONSOLIDATION
*****************************************************************}

procedure TOF_CONSOLIDATION.AnalyseRepertoire(Rep, Mask: string; TOB_temp: TOB);
var TSR: TsearchRec;
  sousmasque, st, path, Site, Event: string;
  Enveloppe: TCollectionEnveloppes;
begin
  st := IncludeTrailingBackslash(Rep);
  path := ExtractFilePath(st);
  repeat
    begin
      { si le masque est composé, on récupère chaque sous-masque }
      sousmasque := ReadTokenSt(Mask);
      if FindFirst(path + sousmasque, faArchive, TSR) = 0 then
        repeat
          Enveloppe := nil;
          try
            Enveloppe := TCollectionEnveloppes.Create(TCollectionEnveloppe);
            Enveloppe.LoadEnveloppe(Path + TSR.Name);
            { Recherche dans l'enveloppe }
            with Enveloppe.Items[0] do
            begin
              if TOB_temp = TOB_DEPART then
                Site := dSite else
                Site := eSite;

              Event := CodeEvent;
              { on affecte que si on ne masque pas OU si on masque et que le site est actif }
              if (not MasquerSitesInactifs) or (MasquerSitesInactifs and SiteActif(Site)) then
              begin
                TOB_temp.PutValue(Site, TOB_temp.GetValue(Site) + 1);
                { on met à jour la tob si la date est plus ancienne }
                if TOB_temp = TOB_DEPART then
                  if DateMsg < StrToDate(TOB_JOURS.GetValue(Site)) then
                    TOB_JOURS.PutValue(Site, DateToStr(DateMsg));
              end;
            end;
          finally
            if Enveloppe <> nil then
              Enveloppe.Free;
          end;
        until FindNext(TSR) <> 0;
      FindClose(TSR);
    end;
  until Mask = '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 22/07/2003
Modifié le ... : 06/08/2003
Description .. : Constitution des TOB : Structure et Valeurs
Suite ........ : Masquage des sites
Mots clefs ... : CONSOLIDATION
*****************************************************************}

procedure TOF_CONSOLIDATION.InitLesTob(iMax: integer);
var i, j, k: integer;
  l: double;
begin
  { pour masquer les sites inactifs et affecter la date par initiale}
  for i := 0 to iMax - 1 do
    if (SiteActif(DIM_SITES.Values[i]) and MasquerSitesInactifs) or (not MasquerSitesInactifs) then
    begin
      TOB_DEPART.AddChampSupValeur(DIM_SITES.Values[i], 0, true);
      TOB_JOURS.PutValue(DIM_SITES.Values[i], DateToStr(iDate2099));
    end;
  { affectation des valeurs }
  AnalyseRepertoire(REP_DEPART, '*.ENV', TOB_DEPART);
  AnalyseRepertoire(REP_TRAITE, '*.ENV', TOB_TRAITE);
  AnalyseRepertoire(REP_REJET, '*.ENV', TOB_REJET_SC);

  { pour masquer les sites sans problème }
  j := TOB_DEPART.ChampsSup.Count - 1;
  k := 0;
  for i := 0 to j do
  begin
    if ((MasquerSitesOK) and (TOB_DEPART.GetValeur(1000 + k) + TOB_REJET_SC.GetValeur(1000 + k) = 0)) then
    begin
      TOB_DEPART.DelChampSup(DIM_SITES.Values[k], true);
      DIM_SITES.Items.Delete(k);
      DIM_SITES.Values.Delete(k);
    end else
      Inc(k);
  end;
  { calcul du nb de jours en erreur }
  j := TOB_JOURS.ChampsSup.Count - 1;
  for i := 0 to j do
  begin
    l := Int(Now - StrToDate(TOB_JOURS.GetValeur(1000 + i)));
    if l < 0 then
      l := 0;
    TOB_JOURS.PutValeur(1000 + i, l);
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 29/07/2003
Modifié le ... : 29/07/2003
Description .. : un site est-il actif ?
Mots clefs ... : CONSOLIDATION
*****************************************************************}

function TOF_CONSOLIDATION.SiteActif(CodeSite: string): boolean;
var LeSite: TCollectionSite;
begin
  result := false;
  LeSite := TX_LesSites.Find(CodeSite);
  if LeSite <> nil then
    result := LeSite.SSI_SITEENABLED;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 30/07/2003
Modifié le ... : 31/07/2003
Description .. : Paramètres de chargement de la tablette des sites
Mots clefs ... : CONSOLIDATION
*****************************************************************}

procedure TOF_CONSOLIDATION.ChargeCombos;
var Where: string;
begin
  Where := '';
  if not Assigned(LesItems) then
    LesItems := HTStringList.Create;
  if not Assigned(LesValeurs) then
    LesValeurs := HTStringList.Create;
  if MasquerSitesInactifs then
    Where := 'SSI_SITEENABLED="X"';
  SQL2Combo('STOXSITES', 'SSI_CODESITE', 'SSI_LIBELLE', Where, 'SSI_CODESITE', LesItems, LesValeurs);
  DIM_SITES.Items := LesItems;
  DIM_SITES.Values := LesValeurs;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 30/07/2003
Modifié le ... : 31/07/2003
Description .. : Chargement de la tablette des sites
Mots clefs ... : CONSOLIDATION
*****************************************************************}

procedure TOF_CONSOLIDATION.SQL2Combo(const NomTable, ChampCode,
  ChampLibelle, Where, OrderBy: string; Items, Valeurs: HTStrings);
var
  SQL: string;
  Q: TQuery;
begin
  if (not Assigned(Items)) or (not Assigned(Valeurs)) then
    Exit;
  Items.Clear;
  Valeurs.Clear;
  try
    SQL := 'SELECT ' + ChampCode + ',' + ChampLibelle + ' FROM ' + NomTable;
    if Where <> '' then
      SQL := SQL + ' WHERE ' + Where;
    if OrderBy <> '' then
      SQL := SQL + ' ORDER BY ' + OrderBy;
    Q := OpenSQL(SQL, True,-1,'',true);
  except
    Exit;
  end;
  Q.First;
  while not Q.EOF do
  begin
    Items.Add(Q.FindField(ChampCode).AsString + ' - ' + Q.FindField(ChampLibelle).AsString);
    Valeurs.Add(Q.FindField(ChampCode).AsString);
    Q.Next;
  end;
  Ferme(Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 30/07/2003
Modifié le ... : 30/07/2003
Description .. : Définition des polices utilisées pour l'affichage
Mots clefs ... : CONSOLIDATION
*****************************************************************}

procedure TOF_CONSOLIDATION.DefineFont;
begin
  ErrorFont := TFont.Create;
  ErrorFont.Name := 'MS SansSerif';
  ErrorFont.Style := [fsBold, fsItalic];
  ErrorFont.Color := clRed;
  GoodFont := TFont.Create;
  GoodFont.Name := 'MS SansSerif';
  GoodFont.Style := [fsBold];
  GoodFont.Color := clGreen;
end;

procedure TOF_CONSOLIDATION.Affichage(TypeAff: boolean = true);
var Valeurs: THDimensionItemList;
  ItemDim: THDimensionItem;
  TitreLig1, TitreCol1, st: string;
  i, j: integer;
  DIM_LIG, DIM_COL: THValCombobox;
begin
  if not CreateTob then
  begin
    PGIInfo('Veuillez cocher au moins une corbeille à afficher');
    exit;
  end;
  SetControlEnabled('BINVERSER', true);
  DisplayMode := TypeAff;
  DefineFont;
  Valeurs := THDimensionItemList.Create;
  MasquerSitesInactifs := TCheckBox(GetControl('CB_INAC')).Checked;
  MasquerSitesOK := TCheckBox(GetControl('CB_PB')).Checked;
  ChargeCombos;

  if DisplayMode then
  begin
    TitreLig1 := TraduireMemoire('Site');
    TitreCol1 := TraduireMemoire('Fichiers');
    DIM_LIG := DIM_SITES;
    DIM_COL := DIM_BIN;
  end else
  begin
    TitreCol1 := TraduireMemoire('Site');
    TitreLig1 := TraduireMemoire('Fichiers');
    DIM_LIG := DIM_BIN;
    DIM_COL := DIM_SITES;
  end;

  InitLesTob(DIM_SITES.Values.Count);

  for i := 0 to DIM_LIG.Values.Count - 1 do
  begin
    st := DIM_LIG.Values[i];
    for j := 0 to DIM_COL.Items.Count - 1 do
    begin
      { création et positionnement sur un nouvelle cellule }
      ItemDim := THDimensionItem.Create;
      ItemDim.Lig1 := st;
      ItemDim.Col1 := DIM_COL.Values[j];

      { affectation de la valeur à l'item }
      if DIM_LIG = DIM_SITES then
        ItemDim.Valeur[1] := TOB_MERE.Detail[j].GetValeur(1000 + i) else
        ItemDim.Valeur[1] := TOB_MERE.Detail[i].GetValeur(1000 + j);

      { affectation de la fonte en fonction de la valeur  }
      if ItemDim.Valeur[1] > 0 then
      begin
        if ((DIM_LIG = DIM_SITES) and (j <> 1)) or ((DIM_LIG = DIM_BIN) and (i <> 1)) then
          ItemDim.ValeurFont[1] := ErrorFont else
          ItemDim.ValeurFont[1] := GoodFont;
      end else
        ItemDim.ValeurFont[1] := GoodFont;

      { ajout de la valeur dans l'objet dimension }
      Valeurs.Add(ItemDim);
    end;
  end;
  CONSO.NbValeurs := 1;
  CONSO.TypeDonnee[1] := dotReel;
  CONSO.Alignment[1] := taCenter;
  CONSO.Formatdonnees.Add(1).ReelDisplayFormat := '#,##0';
  CONSO.ReadOnly[1] := true;
  CONSO.DefaultColWidth := THSpinEdit(GetControl('LARGEUR_COLONNES')).Value;
  CONSO.DefaultCol1Width := THSpinEdit(GetControl('LARGEUR_LIGNES')).Value;
  CONSO.InitDimension('', TitreLig1, '', TitreCol1, '', nil, DIM_LIG, nil, DIM_COL, nil, Valeurs);
  FreeAndNil(Valeurs);
end;

procedure AGLAffichage(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_CONSOLIDATION) then
    TOF_CONSOLIDATION(TOTOF).Affichage(boolean(parms[1])) else
    exit;
end;

function TOF_CONSOLIDATION.CreateTob: boolean;
begin
  result := true;
  { création des tob pour affichage ultérieur }
  FreeAndNil(TOB_MERE);
  if (CD.Checked or CT.Checked or CRC.Checked or CRB.Checked) then
  begin
    TOB_MERE := Tob.Create('', nil, -1);
    TOB_DEPART := Tob.Create('', TOB_MERE, -1);
    TOB_TRAITE := Tob.Create('', TOB_MERE, -1);
    TOB_REJET_SC := Tob.Create('', TOB_MERE, -1);
    TOB_REJET_BTQ := Tob.Create('', TOB_MERE, -1);
    TOB_JOURS := Tob.Create('', TOB_MERE, -1);
  end else
    result := false;
end;

initialization
  registerclasses([TOF_CONSOLIDATION]);
  RegisterAglProc('Affichage', true, 2, AGLAffichage);
end.
