{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : MBOTARIFPXVENTE ()
Mots clefs ... : TOF;MBOTARIFPXVENTE
*****************************************************************}
unit UTofMBOTARIFPXVENTE;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, dbTables, ComCtrls,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  db, FE_Main,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTob, TarifUtil, SaisUtil, vierge, graphics, grids, M3FP,
  UtilArticle, AGLInit;

type
  TOF_MBOTARIFPXVENTE = class(TOF)
  private
    PrixVente: THGRID;
    TOBTarif, TobPxVente, TobPxVenteLigne: Tob;
    ColDepot, ColTarif, colDev, colPrix, colDim: Integer;
    LesColonnes, CodeArticle, TarifArticle, CodeEtab: string;
    PxVenteArt: Double;
    DateTarif: TDateTime;

    procedure GSDoubleClick(Sender: TObject);
    procedure RemplirTobPxVente;
    procedure BModifClick;
    function TarifalaDim(Depot, TarifTypeDepot: string): string;

  public
    Action: TActionFiche;

    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure CodeGras(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

  end;

implementation

procedure TOF_MBOTARIFPXVENTE.OnUpdate;
begin
  inherited;
end;

procedure TOF_MBOTARIFPXVENTE.OnLoad;
begin
  inherited;
  TobPxVente := TOB.Create('les prix', nil, -1);
  DateTarif := StrToDate(GetControlText('DATE'));
  TOBTarif := CreerTobTarifArt(CodeArticle, TarifArticle, '', '', '', '', 'VEN', DateTarif, False, False);
  RemplirTobPxVente;
  TobPxVente.PutGridDetail(PrixVente, False, False, LesColonnes, True);
end;

procedure TOF_MBOTARIFPXVENTE.OnArgument(Arguments: string);
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
        if ChampMul = 'PRIXVENTE' then PxVenteArt := StrToFloat(ValMul);
        if ChampMul = 'DATE' then SetControlText('DATE', ValMul);
        if ChampMul = 'ETABLISSEMENT' then CodeEtab := ValMul;
      end;
    end;
  until Critere = '';
  Ecran.Caption := 'Prix de vente de l''article: ' + Copy(CodeArticle, 1, 18);
  LesColonnes := '_DEPOT;_TARIF;_DEV;_PRIX;_DIM';
  PrixVente := THGRID(GetControl('PRIXVENTE'));
  PrixVente.OnDblClick := GSDoubleClick;
  PrixVente.GetCellCanvas := CodeGras;
  PrixVente.ColCount := 1;
  i := 0;
  S := LesColonnes;
  ColDepot := -1;
  colDev := -1;
  colPrix := -1;
  colDim := -1;
  repeat
    NomCol := ReadTokenSt(S);
    if NomCol <> '' then
    begin
      if NomCol = '_DEPOT' then
      begin
        if i <> 0 then PrixVente.ColCount := PrixVente.ColCount + 1;
        ColDepot := i;
        PrixVente.ColWidths[ColDepot] := 120;
      end
      else
        if NomCol = '_TARIF' then
      begin
        if i <> 0 then PrixVente.ColCount := PrixVente.ColCount + 1;
        ColTarif := i;
        PrixVente.ColWidths[ColTarif] := 100;
      end
      else
        if NomCol = '_DEV' then
      begin
        if i <> 0 then PrixVente.ColCount := PrixVente.ColCount + 1;
        colDev := i;
        PrixVente.ColWidths[colDev] := 70;
        PrixVente.ColAligns[colDev] := taCenter;
      end
      else
        if NomCol = '_PRIX' then
      begin
        if i <> 0 then PrixVente.ColCount := PrixVente.ColCount + 1;
        colPrix := i;
        PrixVente.ColWidths[colPrix] := 70;
        PrixVente.ColLengths[colPrix] := -1;
        PrixVente.ColAligns[colPrix] := taRightJustify;
      end
      else
        if NomCol = '_DIM' then
      begin
        if i <> 0 then PrixVente.colCount := PrixVente.ColCount + 1;
        colDim := i;
        PrixVente.ColWidths[colDim] := 90;
        PrixVente.ColAligns[colDim] := taCenter;
      end;
      Inc(i);
    end;
  until ((St = '') or (NomCol = ''));
  AffecteGrid(PrixVente, Action);
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(PrixVente);
end;

procedure TOF_MBOTARIFPXVENTE.BModifClick;
begin
  TOBTarif.Free;
  TOBTarif := nil;
  TobPxVente.ClearDetail;
  DateTarif := StrToDate(GetControlText('DATE'));
  TOBTarif := CreerTobTarifArt(CodeArticle, TarifArticle, '', '', '', '', 'VEN', DateTarif, False, False);
  RemplirTobPxVente;
  TobPxVente.PutGridDetail(PrixVente, False, False, LesColonnes, True);
end;

procedure TOF_MBOTARIFPXVENTE.OnClose;
begin
  inherited;
  TobPxVenteLigne.Free;
  TobPxVenteLigne := nil;
  TobPxVente.Free;
  TobPxVente := nil;
  TOBTarif.Free;
  TOBTarif := nil;
end;

procedure TOF_MBOTARIFPXVENTE.GSDoubleClick(Sender: TObject);
var Depot: string;
begin
  Depot := PrixVente.Cells[ColDepot, PrixVente.Row];
  if (Depot = '') or (PrixVente.Cells[ColTarif, PrixVente.Row] = '') or (PrixVente.Cells[ColDev, PrixVente.Row] = '') then exit;
  AGLLanceFiche('MBO', 'JUSTIFPXVENTE', '', '', 'CODEARTICLE=' + CodeArticle + ';TARIFARTICLE=' + TarifArticle + ';PRIXVENTEART=' + FloatToStr(PxVenteArt) +
    ';DEPOT=' + Depot + ';DATE=' + DateToStr(DateTarif) + ';DIM=' + PrixVente.Cells[ColDim, PrixVente.Row]);
  BModifClick;
end;

procedure TOF_MBOTARIFPXVENTE.RemplirTobPxVente;
var QEtab: TQuery;
  Prix: Double;
  Dim, SQL: string;
  TOBInfo: TOB;
begin
  TOBInfo := nil;
  SQL := 'Select ET_TYPETARIF,ET_ETABLISSEMENT, ET_LIBELLE, ET_DEVISE from Etabliss';
  if CodeEtab <> '' then SQL := SQL + ' WHERE ET_ETABLISSEMENT="'+ CodeEtab +'"';
  QEtab := OpenSQL(SQL, True);
  while not QEtab.EOF do
  begin
    TobPxVenteLigne := Tob.create('un tarif', TOBPxVente, -1);
    TobPxVenteLigne.AddChampSup('_DEPOT', False);
    TobPxVenteLigne.PutValue('_DEPOT', QEtab.Findfield('ET_LIBELLE').AsString);
    TobPxVenteLigne.AddChampSup('_TARIF', False);
    TobPxVenteLigne.PutValue('_TARIF', RechDom('GCTARIFTYPE1VTE', QEtab.Findfield('ET_TYPETARIF').AsString, False));
    TobPxVenteLigne.AddChampSup('_DEV', False);
    TobPxVenteLigne.PutValue('_DEV', QEtab.Findfield('ET_DEVISE').AsString);
    // Recup Prix
    TobPxVenteLigne.AddChampSup('_PRIX', False);
    TobPxVenteLigne.AddChampSup('_DIM', False);
    Prix := CalculPrixVente(TOBTarif, CodeArticle, TarifArticle, QEtab.FindField('ET_ETABLISSEMENT').AsString, QEtab.Findfield('ET_DEVISE').AsString,
      QEtab.Findfield('ET_TYPETARIF').AsString, PxVenteArt, True);
    TobPxVenteLigne.PutValue('_PRIX', Prix);
    // Tarif du depot ou générique
    if TheTOB <> nil then TOBInfo := TheTOB;
    TheTOB := nil;
    if TOBInfo.GetValue('ISDEPOT') = True then Dim := TarifalaDim(QEtab.FindField('ET_ETABLISSEMENT').AsString, QEtab.FindField('ET_TYPETARIF').AsString)
    else Dim := TarifalaDim('', QEtab.FindField('ET_TYPETARIF').AsString);
    TobPxVenteLigne.PutValue('_DIM', Dim);
    PrixVente.RowCount := PrixVente.RowCount + 1;
    QEtab.next;
  end;
  ferme(QEtab);
  if TOBInfo <> nil then
  begin
    TOBInfo.Free;
  end;
end;

function TOF_MBOTARIFPXVENTE.TarifalaDim(Depot, TarifTypeDepot: string): string;
var i: Integer;
  TobT: Tob;
begin
  Result := '-';
  for i := 0 to TOBTarif.Detail.count - 1 do
  begin
    TobT := TOBTarif.Detail[i];
    if (TobT.GetValue('GF_ARTICLE') <> CodeArticle) and (TobT.GetValue('GF_ARTICLE') <> '') then
      if (TobT.GetValue('GF_DEPOT') = Depot) then
      begin
        if ExisteSQL('Select GFM_PERTARIF from TARIFMODE where GFM_TARFMODE="' + IntToStr(TobT.GetValue('GF_TARFMODE')) + '" and GFM_TYPETARIF="' +
          TarifTypeDepot + '"') then
          if IsCodeArticleUnique(TobT.GetValue('GF_ARTICLE')) then
          begin
            result := 'X';
            Exit;
          end;
      end;
  end;
end;

procedure TOF_MBOTARIFPXVENTE.CodeGras(ACol, ARow: Longint; Canvas: TCanvas;
  AState: TGridDrawState);
begin
  if (ACol = colPrix) and (ARow > 0) then
  begin
    Canvas.Font.Style := [fsBold];
  end;
end;

procedure AGLBMODIFClick(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOTARIFPXVENTE)
    then TOF_MBOTARIFPXVENTE(TOTOF).BModifClick
  else exit;
end;

initialization
  registerclasses([TOF_MBOTARIFPXVENTE]);
  RegisterAglProc('BMODIFClick', TRUE, 0, AGLBMODIFClick);
end.
