unit UTofMBOTarif_Mul;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, HDimension, Entgc, Windows, Utob, AGLInit, ParamSoc,
  {$IFDEF EAGLCLIENT}
  Maineagl, Grids, emul,
  {$ELSE}
  dbctrls, db, dbTables, Fe_Main, mul, DBGrids,
  {$ENDIF}
  M3FP, UtilDimArticle, TarifTiersMode, TarifArticleMode;

type
  TOF_MBOTarif_Mul = class(TOF)
  private
    NatureType: string;
    Grid: THGrid;
    procedure OnArgument(St: string); override;
    procedure OnLoad; override;
    procedure OnVisualiseTarif(CodeTarif, CodeArticle, TarifArticle, CodeTiers, TarifTiers: string);
    procedure InitEntete;
    procedure OnChangeCombo(Quoi: string);

  end;

implementation

procedure TOF_MBOTarif_Mul.OnArgument(St: string);
var Arguments, Critere, ChampMul, ValMul: string;
  i: Integer;
begin
  inherited;
  Arguments := St;
  repeat
    Critere := Trim(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      i := pos('=', Critere);
      if i <> 0 then
      begin
        ChampMul := copy(Critere, 1, i - 1);
        ValMul := copy(Critere, i + 1, length(Critere));
        if ChampMul = 'TYPE' then NatureType := ValMul;
      end;
    end;
  until Critere = '';
  if NatureType = 'ACH' then THValComboBox(GetControl('GFM_TYPETARIF')).DataType := 'GCTARIFTYPE1ACH';
  InitEntete;
end;

procedure TOF_MBOTarif_Mul.OnLoad;
begin
  inherited;
  Grid := THGrid(TFmul(Ecran).FListe);
  if NatureType = 'VTE' then
  begin
    SetControlText('GF_REGIMEPRIX', 'TTC');
    TFmul(Ecran).Caption := 'Consultation - Modification des tarifs de vente';
  end else
  begin
    SetControlText('GF_REGIMEPRIX', 'HT');
    TFmul(Ecran).Caption := 'Consultation - Modification des tarifs d''achat';
  end;
  UpdateCaption(Ecran);
  SetControlText('XX_WHERE', 'GA_STATUTART="GEN" or GA_STATUTART="UNI" or GF_ARTICLE is null or GF_ARTICLE=""');
end;

procedure TOF_MBOTarif_Mul.InitEntete;
var QTyp, QPer: TQuery;
begin
  QTyp := OpenSQL('Select GFT_DEVISE,GFT_CODETYPE,GFT_ETABLISREF from TarifTypMode where GFT_DERUTILISE="X" AND GFT_NATURETYPE="' + NatureType + '"', True);
  if not QTyp.EOF then
  begin
    if QTyp.FindField('GFT_CODETYPE').AsString = '...' then
    begin
      THValComboBox(GetControl('GFM_TYPETARIF')).Value := '...';
      THValComboBox(GetControl('GF_DEVISE')).Value := GetParamSoc('SO_DEVISEPRINC');
      THValComboBox(GetControl('GF_DEPOT')).Value := '' //GetParamSoc('SO_GCDEPOTDEFAUT') ;
    end else
    begin
      THValComboBox(GetControl('GFM_TYPETARIF')).Value := QTyp.FindField('GFT_CODETYPE').AsString;
      THValComboBox(GetControl('GF_DEVISE')).Value := QTyp.FindField('GFT_DEVISE').AsString;
      THValComboBox(GetControl('GF_DEPOT')).Value := QTyp.FindField('GFT_ETABLISREF').AsString;
    end ;
  end  ;
  QPer := OpenSQL('Select * from TarifPer where GFP_DERUTILISE="X"', True);
  if not QPer.EOF then
  begin
    THValComboBox(GetControl('GFM_PERTARIF')).Value := QPer.FindField('GFP_CODEPERIODE').AsString;
    THEdit(GetControl('GF_DATEDEBUT')).Text := QPer.FindField('GFP_DATEDEBUT').AsString;
    THEdit(GetControl('GF_DATEFIN')).Text := QPer.FindField('GFP_DATEFIN').AsString;
  end;
  Ferme(QTyp);
  Ferme(QPer);
end;

procedure TOF_MBOTarif_Mul.OnChangeCombo(Quoi: string);
var Q: TQuery;
  Depot: string;
begin
  if Quoi = 'type' then
  begin
    Q := OpenSQL('Select GFT_DEVISE from TarifTypMode where GFT_CODETYPE="' + GetControlText('GFM_TYPETARIF') + '" AND GFT_NATURETYPE="' + NatureType + '"',
      True);
    THValComboBox(GetControl('GF_DEVISE')).Value := Q.FindField('GFT_DEVISE').AsString;
    Ferme(Q);
    if NatureType = 'VTE' then THValComboBox(GetControl('GF_DEPOT')).Plus := '(ET_TYPETARIF = "' + GetControlText('GFM_TYPETARIF') + '")'
    else THValComboBox(GetControl('GF_DEPOT')).Plus := '(ET_TYPETARIFACH = "' + GetControlText('GFM_TYPETARIF') + '")'
  end
  else
    if Quoi = 'periode' then
  begin
    Q := OpenSQL('Select GFP_DATEDEBUT, GFP_DATEFIN from TarifPer where GFP_CODEPERIODE="' + GetControlText('GFM_PERTARIF') + '"', True);
    if not Q.EOF then
    begin
      THEdit(GetControl('GF_DATEDEBUT')).Text := Q.FindField('GFP_DATEDEBUT').AsString;
      THEdit(GetControl('GF_DATEFIN')).Text := Q.FindField('GFP_DATEFIN').AsString;
    end else
    begin
      THEdit(GetControl('GF_DATEDEBUT')).Text := DateToStr(iDate1900);
      THEdit(GetControl('GF_DATEFIN')).Text := DateToStr(iDate2099);
    end;
    Ferme(Q);
  end
  else
    if Quoi = 'depot' then
  begin
    if GetControlText('GF_DEPOT') = '' then Depot := '...' else Depot := GetControlText('GF_DEPOT');
    Q := OpenSQL('Select GFP_DATEDEBUT, GFP_DATEFIN from TarifPer where GFP_CODEPERIODE="' + GetControlText('GFM_PERTARIF') + '" and GFP_ETABLISSEMENT="' +
      Depot + '"', True);
    if not Q.EOF then
    begin
      THEdit(GetControl('GF_DATEDEBUT')).Text := Q.FindField('GFP_DATEDEBUT').AsString;
      THEdit(GetControl('GF_DATEFIN')).Text := Q.FindField('GFP_DATEFIN').AsString;
    end;
    Ferme(Q);
  end;
end;

procedure TOF_MBOTarif_Mul.OnVisualiseTarif(CodeTarif, CodeArticle, TarifArticle, CodeTiers, TarifTiers: string);
var tobTemp: TOB;
begin
  tobTemp := nil;
  if (CodeArticle <> '') or (TarifArticle <> '') and (Codetiers = '') and (TarifTiers = '') then
  begin
    TobTemp := TOB.Create('_INIT', nil, -1);
    TobTemp.AddChampSup('_CodeTarif', False);
    TobTemp.AddChampSup('_CodeArticle', False);
    TobTemp.AddChampSup('_TarifArticle', False);
    TobTemp.PutValue('_CodeTarif', CodeTarif);
    TobTemp.PutValue('_CodeArticle', CodeArticle);
    TobTemp.PutValue('_TarifArticle', TarifArticle);
    TheTob := TobTemp;
    if NatureType = 'VTE' then EntreeTarifArticleMode(taModif, TRUE)
    else EntreeTarifArticleMode(taModif);
  end;
  if (CodeTiers <> '') or (TarifTiers <> '') then
  begin
    TobTemp := TOB.Create('_INIT', nil, -1);
    TobTemp.AddChampSup('_CodeTarif', False);
    TobTemp.AddChampSup('_CodeTiers', False);
    TobTemp.AddChampSup('_TarifTiers', False);
    TobTemp.PutValue('_CodeTarif', CodeTarif);
    TobTemp.PutValue('_CodeTiers', CodeTiers);
    TobTemp.PutValue('_TarifTiers', TarifTiers);
    TheTob := TobTemp;
    if NatureType = 'VTE' then EntreeTarifTiersMode(taModif, TRUE) ;
  end;
  TobTemp.Free;
end;

procedure AGLOnVisualiseTarif(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFMul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_MBOTarif_Mul) then TOF_MBOTarif_Mul(TOTOF).OnVisualiseTarif(parms[1], parms[2], parms[3], parms[4], parms[5]);
end;

procedure AGLOnChangeCombo(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFMul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_MBOTarif_Mul) then TOF_MBOTarif_Mul(TOTOF).OnChangeCombo(Parms[1]);
end;

initialization
  registerclasses([TOF_MBOTarif_Mul]);
  RegisterAglProc('OnVisualiseTarif', TRUE, 5, AGLOnVisualiseTarif);
  RegisterAglProc('OnChangeCombo', TRUE, 1, AGLOnChangeCombo);
end.
