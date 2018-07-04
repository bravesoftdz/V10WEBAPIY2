{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : JUSTIFPXVENTE ()
Mots clefs ... : TOF;JUSTIFPXVENTE
*****************************************************************}
unit JUSTIFPXVENTE_TOF;

interface

uses StdCtrls,
  Controls,
  Classes, Grids, vierge, SaisUtil, TarifUtil, AGLInit, Graphics, Windows,
  {$IFNDEF EAGLCLIENT}
  db,
  dbtables,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF, UTOB, TarifArticleMode;

type
  TOF_JUSTIFPXVENTE = class(TOF)
  private
    JustifPxVente: THGRID;
    colDepot, colProvenance, colPeriode, colPourcent, colPrix, colArrondi, colDevise, colTarif, colDebut, colFin: Integer;
    LesColonnes, CodeArticle, TarifArticle, Depot: string;
    PxVenteArt: Double;
    DateTarif: TDateTime;
    TOBGrid, TOBTarif, TOBEtab, TOBMode: TOB;

    procedure RemplirTOBGrid;
    procedure ChargeLigneArticle;
    procedure ChargeLigneTarifBase;
    procedure ChargeLigneTarifPromo;
    procedure ChargeLigneTarifRemise;
    function TraiterRemiseCascade(TOBTarif: TOB; IsArticle, IsDepot: Boolean; Code: string; CodeTarif: Integer): Integer;
    procedure ChargeLigneRemCascade(Remise: Double; Arrondi, Depot: string; Tarfmode, CodeTarif: Integer; IsDepot: Boolean);
    procedure LoadLesTob;
    function RecupDatePeriode(Periode: string; IsDepot: Boolean): string;
    // Fonction du grid
    procedure GSDoubleClick(Sender: TObject);
    procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

  public
    Action: TActionFiche;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;

implementation

procedure TOF_JUSTIFPXVENTE.OnNew;
begin
  inherited;
end;

procedure TOF_JUSTIFPXVENTE.OnDelete;
begin
  inherited;
end;

procedure TOF_JUSTIFPXVENTE.OnUpdate;
begin
  inherited;
end;

procedure TOF_JUSTIFPXVENTE.OnArgument(Arguments: string);
var St, S, NomCol, Critere, ChampMul, ValMul: string;
  x, i: Integer;
begin
  Action := taConsult;
  St := Arguments;
  repeat
    Critere := Trim(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'CODEARTICLE' then CodeArticle := ValMul;
        if ChampMul = 'TARIFARTICLE' then TarifArticle := ValMul;
        if ChampMul = 'PRIXVENTEART' then PxVenteArt := StrToFloat(ValMul);
        if ChampMul = 'DEPOT' then Depot := ValMul;
        if ChampMul = 'DATE' then DateTarif := StrToDate(ValMul);
      end;
    end;
  until Critere = '';
  Ecran.Caption := 'Justification du prix vente de l''article: ' + Copy(CodeArticle, 1, 18);
  LesColonnes := '_DEPOT;_PROVENANCE;_PERIODE;_DEBUT;_FIN;_POURCENT;_PRIX;_ARRONDI;_DEVISE;_TARIF';
  JustifPxVente := THGRID(GetControl('JUSTIFPRIX'));
  JustifPxVente.OnDblClick := GSDoubleClick;
  JustifPxVente.GetCellCanvas := GetCellCanvas;
  JustifPxVente.ColCount := 1;
  i := 0;
  S := LesColonnes;
  colDepot := -1;
  colProvenance := -1;
  colPeriode := -1;
  colPourcent := -1;
  colPrix := -1;
  colArrondi := -1;
  colDevise := -1;
  colTarif := -1;
  colDebut := -1;
  colFin := -1;
  repeat
    NomCol := ReadTokenSt(S);
    if NomCol <> '' then
    begin
      if NomCol = '_DEPOT' then
      begin
        if i <> 0 then JustifPxVente.ColCount := JustifPxVente.ColCount + 1;
        ColDepot := i;
        JustifPxVente.ColWidths[ColDepot] := 40;
      end
      else
        if NomCol = '_PROVENANCE' then
      begin
        if i <> 0 then JustifPxVente.colCount := JustifPxVente.ColCount + 1;
        colProvenance := i;
        JustifPxVente.ColWidths[colProvenance] := 95;
      end
      else
        if NomCol = '_PERIODE' then
      begin
        if i <> 0 then JustifPxVente.colCount := JustifPxVente.ColCount + 1;
        colPeriode := i;
        JustifPxVente.ColWidths[colPeriode] := 113;
      end
      else
        if NomCol = '_DEBUT' then
      begin
        if i <> 0 then JustifPxVente.colCount := JustifPxVente.ColCount + 1;
        colDebut := i;
        JustifPxVente.ColWidths[colDebut] := 45;
        JustifPxVente.ColAligns[colDebut] := taLeftJustify;
      end
      else
        if NomCol = '_FIN' then
      begin
        if i <> 0 then JustifPxVente.colCount := JustifPxVente.ColCount + 1;
        colFin := i;
        JustifPxVente.ColWidths[colFin] := 45;
        JustifPxVente.ColAligns[colFin] := taLeftJustify;
      end
      else
        if NomCol = '_POURCENT' then
      begin
        if i <> 0 then JustifPxVente.colCount := JustifPxVente.ColCount + 1;
        ColPourcent := i;
        JustifPxVente.ColWidths[ColPourcent] := 30;
        JustifPxVente.ColAligns[ColPourcent] := taRightJustify;
      end
      else
        if NomCol = '_PRIX' then
      begin
        if i <> 0 then JustifPxVente.ColCount := JustifPxVente.ColCount + 1;
        colPrix := i;
        JustifPxVente.ColWidths[colPrix] := 43;
        JustifPxVente.ColLengths[colPrix] := -1;
        JustifPxVente.ColAligns[colPrix] := taRightJustify;
      end
      else
        if NomCol = '_ARRONDI' then
      begin
        if i <> 0 then JustifPxVente.ColCount := JustifPxVente.ColCount + 1;
        colArrondi := i;
        JustifPxVente.ColWidths[colArrondi] := 0;
        JustifPxVente.ColLengths[colArrondi] := -1;
      end
      else
        if NomCol = '_DEVISE' then
      begin
        if i <> 0 then JustifPxVente.ColCount := JustifPxVente.ColCount + 1;
        colDevise := i;
        JustifPxVente.ColWidths[colDevise] := 35;
        JustifPxVente.ColAligns[colDevise] := taCenter;
      end
      else
        if NomCol = '_TARIF' then
      begin
        if i <> 0 then JustifPxVente.ColCount := JustifPxVente.ColCount + 1;
        colTarif := i;
        JustifPxVente.ColWidths[colTarif] := 0;
        JustifPxVente.ColLengths[colTarif] := -1;
      end;
      Inc(i);
    end;
  until ((St = '') or (NomCol = ''));
  AffecteGrid(JustifPxVente, Action);
end;

procedure TOF_JUSTIFPXVENTE.OnLoad;
begin
  inherited;
  // Créer les Tob
  TOBEtab := TOB.Create('ETAB', nil, -1);
  TOBMode := TOB.Create('', nil, -1);
  // Charger les Tob
  LoadLesTob;
  // Chargement Grid
  TOBGrid := TOB.Create('les lignes', nil, -1);
  RemplirTOBGrid;
  TOBGrid.PutGridDetail(JustifPxVente, False, False, LesColonnes, True);
end;

procedure TOF_JustifPxVente.OnClose;
begin
  inherited;
  TOBGrid.Free;
  TOBGrid := nil;
  TOBTarif.Free;
  TOBTarif := nil;
  TOBEtab.Free;
  TOBEtab := nil;
  TOBMode.Free;
  TOBMode := nil;
end;

procedure TOF_JustifPxVente.RemplirTOBGrid;
begin
  ChargeLigneArticle;
  ChargeLigneTarifBase;
  ChargeLigneTarifPromo;
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(JustifPxVente);
end;

// Charge le prix article

procedure TOF_JustifPxVente.ChargeLigneArticle;
var TOBGridLigne, TOBM: TOB;
  TypeTarif: string;
begin
  // Remplir TOBGridLigne
  TOBGridLigne := Tob.create('une ligne', TOBGrid, -1);
  TOBGridLigne.AddChampSup('_DEPOT', False);
  TOBGridLigne.AddChampSup('_PROVENANCE', False);
  TOBGridLigne.AddChampSup('_PERIODE', False);
  TOBGridLigne.AddChampSup('_DEBUT', False);
  TOBGridLigne.AddChampSup('_FIN', False);
  TOBGridLigne.AddChampSup('_POURCENT', False);
  TOBGridLigne.AddChampSup('_PRIX', False);
  TOBGridLigne.AddChampSup('_ARRONDI', False);
  TOBGridLigne.AddChampSup('_DEVISE', False);
  TOBGridLigne.AddChampSup('_TARIF', False);
  TOBGridLigne.PutValue('_DEPOT', '...');
  TOBGridLigne.PutValue('_PROVENANCE', 'Prix fiche article');
  TOBGridLigne.PutValue('_PRIX', PxVenteArt);
  TOBGridLigne.PutValue('_POURCENT', '');
  TOBGridLigne.PutValue('_DEVISE', V_PGI.DevisePivot);
  TOBGridLigne.PutValue('_TARIF', 0);
  // Conversion prix article en fonction du type de tarif
  TOBM := TOBMode.FindFirst(['GFM_DEVISE', 'GFM_TYPETARIF'], [TOBEtab.GetValue('ET_DEVISE'), TOBEtab.GetValue('ET_TYPETARIF')], False);
  if TOBM <> nil then
  begin
    PxVenteArt := PxVenteArt * TOBM.GetValue('GFM_COEF');
    if TOBM.GetValue('GFM_COEF') <> 1 then
    begin
      TypeTarif := RechDom('GCTARIFTYPE1VTE', TOBM.GetValue('GFM_TypeTarif'), False);
      TOBGridLigne := Tob.create('une ligne', TOBGrid, -1);
      TOBGridLigne.AddChampSup('_DEPOT', False);
      TOBGridLigne.AddChampSup('_PROVENANCE', False);
      TOBGridLigne.AddChampSup('_PERIODE', False);
      TOBGridLigne.AddChampSup('_POURCENT', False);
      TOBGridLigne.AddChampSup('_PRIX', False);
      TOBGridLigne.AddChampSup('_ARRONDI', False);
      TOBGridLigne.AddChampSup('_DEVISE', False);
      TOBGridLigne.AddChampSup('_TARIF', False);
      TOBGridLigne.PutValue('_DEPOT', '...');
      TOBGridLigne.PutValue('_PROVENANCE', 'Conversion ' + TypeTarif + ' (*' + FloatToStr(TOBM.GetValue('GFM_COEF')) + ')');
      TOBGridLigne.PutValue('_PRIX', PxVenteArt);
      TOBGridLigne.PutValue('_POURCENT', '');
      TOBGridLigne.PutValue('_DEVISE', TOBEtab.GetValue('ET_DEVISE'));
      TOBGridLigne.PutValue('_TARIF', 0);
    end;
  end;
end;

// Charge le tarif de base

procedure TOF_JustifPxVente.ChargeLigneTarifBase;
var TOBGridLigne, TOBInfo, TOBM: TOB;
  PrixBase: Double;
  TypeTarif, Periode, Date: string;
begin
  TOBInfo := nil;
  // Recup Prix
  PrixBase := RechPrixTarifBase(TOBTarif, TOBMode, CodeArticle, TOBEtab.GetValue('ET_ETABLISSEMENT'), TOBEtab.GetValue('ET_DEVISE'),
    TOBEtab.GetValue('ET_TYPETARIF'), 'VTE', True);
  if PrixBase = 0 then exit; // Pas de chargement si pas de prix
  // Recup info Type et période
  if TheTOB <> nil then TOBInfo := TheTOB;
  TheTOB := nil;
  TOBM := TOBMode.FindFirst(['GFM_TARFMODE'], [TOBInfo.GetValue('TARFMODE')], False);
  if TOBM <> nil then
  begin
    Periode := TOBM.GetValue('GFM_PERTARIF');
    TypeTarif := TOBM.GetValue('GFM_TYPETARIF');
  end;
  Date := RecupDatePeriode(Periode, TOBInfo.GetValue('ISDEPOT'));
  Periode := RechDom('GCTARIFPERIODE1', Periode, False);
  TypeTarif := RechDom('GCTARIFTYPE1VTE', TypeTarif, False);
  // Remplir TOBGridLigne
  TOBGridLigne := Tob.create('une ligne', TOBGrid, -1);
  TOBGridLigne.AddChampSup('_DEPOT', False);
  TOBGridLigne.AddChampSup('_PROVENANCE', False);
  TOBGridLigne.AddChampSup('_PERIODE', False);
  TOBGridLigne.AddChampSup('_DEBUT', False);
  TOBGridLigne.AddChampSup('_FIN', False);
  TOBGridLigne.AddChampSup('_POURCENT', False);
  TOBGridLigne.AddChampSup('_PRIX', False);
  TOBGridLigne.AddChampSup('_ARRONDI', False);
  TOBGridLigne.AddChampSup('_DEVISE', False);
  TOBGridLigne.AddChampSup('_TARIF', False);
  if TOBInfo.GetValue('ISDEPOT') = True then TOBGridLigne.PutValue('_DEPOT', TOBEtab.GetValue('ET_LIBELLE'))
  else TOBGridLigne.PutValue('_DEPOT', '...');
  TOBGridLigne.PutValue('_PROVENANCE', TypeTarif);
  TOBGridLigne.PutValue('_PERIODE', Periode);
  TOBGridLigne.PutValue('_DEBUT', ReadTokenSt(Date));
  TOBGridLigne.PutValue('_FIN', ReadTokenSt(Date));
  TOBGridLigne.PutValue('_PRIX', PrixBase);
  TOBGridLigne.PutValue('_POURCENT', '');
  TOBGridLigne.PutValue('_DEVISE', TOBEtab.GetValue('ET_DEVISE'));
  TOBGridLigne.PutValue('_TARIF', TOBInfo.GetValue('CODETARIF'));
  if TOBInfo <> nil then
  begin
    TOBInfo.free;
  end;
end;

// Charge les lignes des tarifs promotions en montants

procedure TOF_JustifPxVente.ChargeLigneTarifPromo;
var TOBGridLigne, TOBInfo, TOBM: TOB;
  PrixSolde: Double;
  TypeTarif, Periode, Info, Date: string;
begin
  TOBInfo := nil;
  TOBM := nil;
  // Recup Prix spécifique
  Info := RechTarifSpec(TOBTarif, TOBMode, CodeArticle, TOBEtab.GetValue('ET_ETABLISSEMENT'), TOBEtab.GetValue('ET_DEVISE'), TOBEtab.GetValue('ET_TYPETARIF'),
    'VTE', True);
  PrixSolde := Valeur(ReadTokenSt(Info));
  // Recup remise
  if PrixSolde = 0 then
  begin
    ChargeLigneTarifRemise;
    exit;
  end;
  // Recup info Type et période
  if TheTOB <> nil then TOBInfo := TheTOB;
  TheTOB := nil;
  if TOBInfo <> nil then TOBM := TOBMode.FindFirst(['GFM_TARFMODE'], [TOBInfo.GetValue('TARFMODE')], False);
  if TOBM <> nil then
  begin
    Periode := TOBM.GetValue('GFM_PERTARIF');
    TypeTarif := TOBM.GetValue('GFM_TYPETARIF');
  end;
  Date := RecupDatePeriode(Periode, TOBInfo.GetValue('ISDEPOT'));
  Periode := RechDom('GCTARIFPERIODE1', Periode, False);
  TypeTarif := RechDom('GCTARIFTYPE1VTE', TypeTarif, False);
  // Remplir TOBGridLigne
  TOBGridLigne := Tob.create('une ligne', TOBGrid, -1);
  TOBGridLigne.AddChampSup('_DEPOT', False);
  TOBGridLigne.AddChampSup('_PROVENANCE', False);
  TOBGridLigne.AddChampSup('_PERIODE', False);
  TOBGridLigne.AddChampSup('_DEBUT', False);
  TOBGridLigne.AddChampSup('_FIN', False);
  TOBGridLigne.AddChampSup('_POURCENT', False);
  TOBGridLigne.AddChampSup('_PRIX', False);
  TOBGridLigne.AddChampSup('_ARRONDI', False);
  TOBGridLigne.AddChampSup('_DEVISE', False);
  TOBGridLigne.AddChampSup('_TARIF', False);
  if TOBInfo.GetValue('ISDEPOT') = True then TOBGridLigne.PutValue('_DEPOT', TOBEtab.GetValue('ET_LIBELLE'))
  else TOBGridLigne.PutValue('_DEPOT', '...');
  TOBGridLigne.PutValue('_PROVENANCE', TypeTarif);
  TOBGridLigne.PutValue('_PERIODE', Periode);
  TOBGridLigne.PutValue('_DEBUT', ReadTokenSt(Date));
  TOBGridLigne.PutValue('_FIN', ReadTokenSt(Date));
  TOBGridLigne.PutValue('_PRIX', PrixSolde);
  TOBGridLigne.PutValue('_POURCENT', '');
  TOBGridLigne.PutValue('_DEVISE', TOBEtab.GetValue('ET_DEVISE'));
  TOBGridLigne.PutValue('_TARIF', TOBInfo.GetValue('CODETARIF'));
  if TOBInfo <> nil then
  begin
    TOBInfo.free;
  end;
end;

// Charge les lignes des tarifs promotions en pourcentages

procedure TOF_JustifPxVente.ChargeLigneTarifRemise;
var TOBGridLigne, TOBInfo, TOBInfoArt, TOBInfoCatArt, TOBM: TOB;
  PrixSolde, RemiseArt, RemiseCatArt, Remise: Double;
  TypeTarif, Periode, InfoCatArt, InfoArt, Arrondi, Code, Date: string;
  IsArticle: Boolean;
  NbLigneCas: Integer;
begin
  TobInfoArt := nil;
  TobInfoCatArt := nil;
  TOBM := nil;
  // Remise Article
  InfoArt := CalcRemArt(TOBTarif, TOBMode, CodeArticle, TOBEtab.GetValue('ET_ETABLISSEMENT'), TOBEtab.GetValue('ET_DEVISE'), TOBEtab.GetValue('ET_TYPETARIF'),
    'VTE', True);
  RemiseArt := Valeur(ReadTokenSt(InfoArt));
  // Recup info Type et période
  if TheTOB <> nil then TOBInfoArt := TheTOB;
  TheTOB := nil;
  // Remise catégorie article
  InfoCatArt := CalcRemCatArt(TOBTarif, TOBMode, TarifArticle, TOBEtab.GetValue('ET_ETABLISSEMENT'), TOBEtab.GetValue('ET_DEVISE'),
    TOBEtab.GetValue('ET_TYPETARIF'), 'VTE', True);
  RemiseCatArt := Valeur(ReadTokenSt(InfoCatArt));
  // Recup info Type et période
  if TheTOB <> nil then TOBInfoCatArt := TheTOB;
  TheTOB := nil;

  if (RemiseArt = 0) and (RemiseCatArt = 0) then exit;
  if RemiseArt >= RemiseCatArt then
  begin
    Arrondi := ReadTokenSt(InfoArt);
    Remise := RemiseArt;
    TobInfo := TobInfoArt;
    IsArticle := True;
    Code := CodeArticle;
  end
  else
  begin
    Arrondi := ReadTokenSt(InfoCatArt);
    Remise := RemiseCatArt;
    TypeTarif := 'Tarif Catégorie Article';
    TobInfo := TobInfoCatArt;
    IsArticle := False;
    Code := TarifArticle;
  end;
  PrixSolde := CalculPrixVente(TOBTarif, CodeArticle, TarifArticle, TOBEtab.GetValue('ET_ETABLISSEMENT'), TOBEtab.GetValue('ET_DEVISE'),
    TOBEtab.GetValue('ET_TYPETARIF'), PxVenteArt, True);
  JustifPxVente.ColWidths[colArrondi] := 75;
  TheTob := nil;
  if TOBInfo <> nil then TOBM := TOBMode.FindFirst(['GFM_TARFMODE'], [TOBInfo.GetValue('TARFMODE')], False);
  if TOBM <> nil then
  begin
    Periode := TOBM.GetValue('GFM_PERTARIF');
    Date := RecupDatePeriode(Periode, TOBInfo.GetValue('ISDEPOT'));
    Periode := RechDom('GCTARIFPERIODE1', Periode, False);
    if TypeTarif = '' then
    begin
      TypeTarif := TOBM.GetValue('GFM_TYPETARIF');
      TypeTarif := RechDom('GCTARIFTYPE1VTE', TypeTarif, False);
    end;
  end;
  if Arrondi <> '' then Arrondi := RechDom('GCCODEARRONDI', Arrondi, False);
  NbLigneCas := TraiterRemiseCascade(TOBTarif, IsArticle, TOBInfo.GetValue('ISDEPOT'), Code, TOBInfo.GetValue('CODETARIF'));
  // Remplir TOBGridLigne
  // Ligne de remise si pas de cascade
  if (NbLigneCas = 0) and (PrixSolde <> 0) then
  begin
    TOBGridLigne := Tob.create('une ligne', TOBGrid, -1);
    TOBGridLigne.AddChampSup('_DEPOT', False);
    TOBGridLigne.AddChampSup('_PROVENANCE', False);
    TOBGridLigne.AddChampSup('_PERIODE', False);
    TOBGridLigne.AddChampSup('_DEBUT', False);
    TOBGridLigne.AddChampSup('_FIN', False);
    TOBGridLigne.AddChampSup('_POURCENT', False);
    TOBGridLigne.AddChampSup('_PRIX', False);
    TOBGridLigne.AddChampSup('_ARRONDI', False);
    TOBGridLigne.AddChampSup('_DEVISE', False);
    TOBGridLigne.AddChampSup('_TARIF', False);
    if TOBInfo.GetValue('ISDEPOT') = True then TOBGridLigne.PutValue('_DEPOT', TOBEtab.GetValue('ET_LIBELLE'))
    else TOBGridLigne.PutValue('_DEPOT', '...');
    TOBGridLigne.PutValue('_PROVENANCE', TypeTarif);
    TOBGridLigne.PutValue('_PERIODE', Periode);
    TOBGridLigne.PutValue('_DEBUT', ReadTokenSt(Date));
    TOBGridLigne.PutValue('_FIN', ReadTokenSt(Date));
    TOBGridLigne.PutValue('_PRIX', PrixSolde);
    TOBGridLigne.PutValue('_ARRONDI', Arrondi);
    TOBGridLigne.PutValue('_POURCENT', Remise);
    TOBGridLigne.PutValue('_DEVISE', TOBEtab.GetValue('ET_DEVISE'));
    TOBGridLigne.PutValue('_TARIF', TOBInfo.GetValue('CODETARIF'));
  end;
  if (NbLigneCas < 2) then
  begin
    TOBGrid.Detail[TOBGrid.Detail.count - 1].PutValue('_PRIX', PrixSolde);
    exit;
  end;
  // Si cascade: ligne calcul cascade
  TOBGridLigne := Tob.create('une ligne', TOBGrid, -1);
  TOBGridLigne.AddChampSup('_DEPOT', False);
  TOBGridLigne.AddChampSup('_PROVENANCE', False);
  TOBGridLigne.AddChampSup('_PERIODE', False);
  TOBGridLigne.AddChampSup('_POURCENT', False);
  TOBGridLigne.AddChampSup('_PRIX', False);
  TOBGridLigne.AddChampSup('_ARRONDI', False);
  TOBGridLigne.AddChampSup('_DEVISE', False);
  TOBGridLigne.AddChampSup('_TARIF', False);
  if TOBInfo.GetValue('ISDEPOT') = True then TOBGridLigne.PutValue('_DEPOT', TOBEtab.GetValue('ET_LIBELLE'))
  else TOBGridLigne.PutValue('_DEPOT', '...');
  TOBGridLigne.PutValue('_PROVENANCE', 'Calcul cascade');
  TOBGridLigne.PutValue('_PERIODE', '');
  TOBGridLigne.PutValue('_PRIX', PrixSolde);
  TOBGridLigne.PutValue('_ARRONDI', Arrondi);
  TOBGridLigne.PutValue('_POURCENT', Remise);
  TOBGridLigne.PutValue('_DEVISE', TOBEtab.GetValue('ET_DEVISE'));
  TOBGridLigne.PutValue('_TARIF', 0);
  if TOBInfo <> nil then
  begin
    TOBInfo.Free;
  end;
  if TOBInfoArt <> nil then
  begin
    TOBInfoArt.free;
  end;
  if TOBInfoCatArt <> nil then
  begin
    TOBInfoCatArt.free;
  end;
end;

// Traiter les lignes de remise

function TOF_JustifPxVente.TraiterRemiseCascade(TOBTarif: TOB; IsArticle, IsDepot: Boolean; Code: string; CodeTarif: Integer): Integer;
var Remise: Double;
  Arrondi: string;
  TarfMode, i, j: integer;
begin
  j := 0;
  for i := 0 to TOBTarif.Detail.count - 1 do
  begin
    if TOBTarif.detail[i].GetValue('GF_REMISE') = 0 then continue
    else
    begin
      if (IsArticle) and (TOBTarif.Detail[i].GetValue('GF_ARTICLE') <> Code) then continue;
      if (not IsArticle) and (TOBTarif.Detail[i].GetValue('GF_TARIFARTICLE') <> Code) then continue;
      if (not IsDepot) and (TOBTarif.Detail[i].GetValue('GF_DEPOT') <> '') then break;
      if (IsDepot) and (TOBTarif.Detail[i].GetValue('GF_DEPOT') = '') then break;
      if TOBTarif.Detail[i].GetValue('GF_CASCADEREMISE') <> 'CAS' then break;
      Remise := TOBTarif.Detail[i].GetValue('GF_REMISE');
      Arrondi := TOBTarif.Detail[i].GetValue('GF_ARRONDI');
      TarfMode := TOBTarif.Detail[i].GetValue('GF_TARFMODE');
      Depot := TOBTarif.Detail[i].GetValue('GF_DEPOT');
      ChargeLigneRemCascade(Remise, Arrondi, Depot, Tarfmode, CodeTarif, IsDepot);
      inc(j);
    end;
  end;
  Result := j;
end;

// Charge les lignes de remises

procedure TOF_JustifPxVente.ChargeLigneRemCascade(Remise: Double; Arrondi, Depot: string; Tarfmode, CodeTarif: Integer; IsDepot: Boolean);
var TOBM, TOBGridLigne: TOB;
  Periode, TypeTarif, Date: string;
begin
  if Arrondi <> '' then Arrondi := RechDom('GCCODEARRONDI', Arrondi, False);
  TOBM := TOBMode.FindFirst(['GFM_TARFMODE'], [Tarfmode], False);
  if TOBM <> nil then
  begin
    Periode := TOBM.GetValue('GFM_PERTARIF');
    Date := RecupDatePeriode(Periode, IsDepot);
    Periode := RechDom('GCTARIFPERIODE1', Periode, False);
    TypeTarif := TOBM.GetValue('GFM_TYPETARIF');
    TypeTarif := RechDom('GCTARIFTYPE1VTE', TypeTarif, False);
  end;
  // Remplir TOBGridLigne
  TOBGridLigne := Tob.create('une ligne', TOBGrid, -1);
  TOBGridLigne.AddChampSup('_DEPOT', False);
  TOBGridLigne.AddChampSup('_PROVENANCE', False);
  TOBGridLigne.AddChampSup('_PERIODE', False);
  TOBGridLigne.AddChampSup('_DEBUT', False);
  TOBGridLigne.AddChampSup('_FIN', False);
  TOBGridLigne.AddChampSup('_POURCENT', False);
  TOBGridLigne.AddChampSup('_PRIX', False);
  TOBGridLigne.AddChampSup('_ARRONDI', False);
  TOBGridLigne.AddChampSup('_DEVISE', False);
  TOBGridLigne.AddChampSup('_TARIF', False);
  if IsDepot then TOBGridLigne.PutValue('_DEPOT', TOBEtab.GetValue('ET_LIBELLE'))
  else TOBGridLigne.PutValue('_DEPOT', '...');
  TOBGridLigne.PutValue('_PROVENANCE', TypeTarif);
  TOBGridLigne.PutValue('_PERIODE', Periode);
  TOBGridLigne.PutValue('_DEBUT', ReadTokenSt(Date));
  TOBGridLigne.PutValue('_FIN', ReadTokenSt(Date));
  TOBGridLigne.PutValue('_PRIX', '');
  TOBGridLigne.PutValue('_ARRONDI', Arrondi);
  TOBGridLigne.PutValue('_POURCENT', Remise);
  TOBGridLigne.PutValue('_DEVISE', TOBEtab.GetValue('ET_DEVISE'));
  TOBGridLigne.PutValue('_TARIF', CodeTarif);
end;

procedure TOF_JustifPxVente.LoadLesTob;
var QEtab, Qmode: TQuery;
begin
  // Chargement info établissement
  QEtab := OpenSQL('Select ET_ETABLISSEMENT,ET_LIBELLE,ET_TYPETARIF,ET_DEVISE from Etabliss where ET_LIBELLE="' + Depot + '"', True);
  TOBEtab.SelectDB('', QEtab);
  TOBEtab.LoadDB;
  // Chargement info tarif
  QMode :=
    OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_PERTARIF,GFM_DATEDEBUT,GFM_PROMO,GFM_COEF,GFM_DEVISE from TARIFMODE where gfm_typetarif in ("'
    + QEtab.FindField('ET_TYPETARIF').AsString + '","...") and gfm_devise="' + QEtab.FindField('ET_DEVISE').AsString +
    '" order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True);
  if not Qmode.EOF then TOBMode.LoadDetailDB('TARIFMODE', '', '', QMode, False);
  Ferme(QMode);
  Ferme(QEtab);
  // Chargement Tarif
  TOBTarif := CreerTobTarifArt(CodeArticle, TarifArticle, '', '', '', Depot, 'VEN', DateTarif, False, False);
end;

// Renvoie les dates de la periode suivant l'etablissement ou non

function TOF_JustifPxVente.RecupDatePeriode(Periode: string; IsDepot: Boolean): string;
var InfoPeriode, Debut, Fin: string;
begin
  if IsDepot then InfoPeriode := RecupInfoPeriode(Periode, TOBEtab.GetValue('ET_ETABLISSEMENT')) else
    InfoPeriode := RecupInfoPeriode(Periode, '');
  Debut := ReadTokenSt(InfoPeriode);
  Fin := ReadTokenSt(InfoPeriode);
  Result := Debut + ';' + Fin;
end;
//********************************************
//************ Fonction du Grid **************
//********************************************

procedure TOF_JustifPxVente.GSDoubleClick(Sender: TObject);
var TOBTemp: TOB;
  TarifAct: TActionFiche;
begin
  if JustifPxVente.Cells[colTarif, JustifPxVente.Row] = '0' then exit;
  TobTemp := TOB.Create('_INIT', nil, -1);
  TobTemp.AddChampSup('_CodeTarif', False);
  TobTemp.AddChampSup('_CodeArticle', False);
  TobTemp.AddChampSup('_TarifArticle', False);
  TobTemp.PutValue('_CodeTarif', JustifPxVente.Cells[colTarif, JustifPxVente.Row]);
  TobTemp.PutValue('_CodeArticle', CodeArticle);
  TobTemp.PutValue('_TarifArticle', TarifArticle);
  TheTob := TobTemp;
  if CtxFO in V_PGI.PGIContexte then TarifAct := taConsult else TarifAct := taModif; //NA : 20/09/2002
  EntreeTarifArticleMode(TarifAct, TRUE);
  TOBGrid.ClearDetail;
  TOBTarif.ClearDetail;
  TOBEtab.ClearDetail;
  TOBMode.ClearDetail;
  LoadLesTob;
  RemplirTOBGrid;
  TOBGrid.PutGridDetail(JustifPxVente, False, False, LesColonnes, True);
  TheTob := nil;
  TobTemp.Free;
end;

procedure TOF_JustifPxVente.GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if (ARow = TOBGrid.Detail.count) then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsBold];
    if ACol = ColPrix then Canvas.Font.Color := clRed;
  end;
  if (ACol = ColPrix) and (JustifPxVente.Cells[ColPrix, Arow] <> '') and (JustifPxVente.Cells[ColPourcent, Arow + 1] <> '') then
    Canvas.Font.Style := Canvas.Font.Style + [fsBold];
  if (ACol = ColPourcent) and (JustifPxVente.Cells[ColPourcent, Arow] <> '') then
    Canvas.Font.Style := Canvas.Font.Style + [fsBold];
end;

initialization
  registerclasses([TOF_JUSTIFPXVENTE]);
end.
