{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TARIFSUPPRESSION ()
Mots clefs ... : TOF;TARIFSUPPRESSION
*****************************************************************}
unit TARIFSUPPRESSION_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db, dbtables, DBGrids, mul,
  {$ELSE}
  eMul,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, ParamSoc, UTob, M3FP;

type
  TOF_TARIFSUPPRESSION = class(TOF)
    Grid: THGrid;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnChangeCombo(Quoi: string);
    procedure InitEntete;
    procedure SuppressionTarif;

  private
    NatureType: string;

  end;

procedure SupprimerTousLesTarifs(CodeTarif: string);
procedure RechercheToutTarif(var TousCodeTarif, CodeArticle, Depot: string; TarfMode: Integer);
procedure AGLOnChangeCombo(parms: array of variant; nb: integer);
procedure AGLSuppressionTarif(parms: array of variant; nb: integer);

implementation

procedure TOF_TARIFSUPPRESSION.OnUpdate;
begin
  inherited;
end;

procedure TOF_TARIFSUPPRESSION.OnLoad;
begin
  inherited;
  Grid := THGrid(TFmul(Ecran).FListe);
  if NatureType = 'VTE' then
  begin
    SetControlText('GF_REGIMEPRIX', 'TTC');
    TFmul(Ecran).Caption := 'Suppression des tarifs de vente';
  end else
  begin
    SetControlText('GF_REGIMEPRIX', 'HT');
    TFmul(Ecran).Caption := 'Suppression des tarifs d''achat';
  end;
  UpdateCaption(Ecran);
  SetControlText('XX_WHERE', 'GA_STATUTART="GEN" or GA_STATUTART="UNI" or GF_ARTICLE is null or GF_ARTICLE=""');
end;

// Mise à jour de l'entête

procedure TOF_TARIFSUPPRESSION.OnChangeCombo(Quoi: string);
var Q: TQuery;
  Depot: string;
begin
  if Quoi = 'type' then
  begin
    Q := OpenSQL('Select GFT_DEVISE from TarifTypMode where GFT_CODETYPE="' + GetControlText('GFM_TYPETARIF') + '"', True);
    THValComboBox(GetControl('GF_DEVISE')).Value := Q.FindField('GFT_DEVISE').AsString;
    Ferme(Q);
    if NatureType = 'VTE' then THValComboBox(GetControl('GF_DEPOT')).Plus := '(ET_TYPETARIF = "' + GetControlText('GFM_TYPETARIF') + '")'
    else THValComboBox(GetControl('GF_DEPOT')).Plus := '(ET_TYPETARIFACH = "' + GetControlText('GFM_TYPETARIF') + '")'
  end
  else if Quoi = 'periode' then
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
  else if Quoi = 'depot' then
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

procedure TOF_TARIFSUPPRESSION.OnArgument(S: string);
var Arguments, Critere, ChampMul, ValMul: string;
  i: Integer;
begin
  inherited;
  Arguments := S;
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
  initEntete;
end;

// Initialisation de l'entête

procedure TOF_TARIFSUPPRESSION.InitEntete;
var QTyp, QPer: TQuery;
begin
  QTyp := OpenSQL('Select GFT_DEVISE,GFT_CODETYPE,GFT_ETABLISREF from TarifTypMode where GFT_DERUTILISE="X" AND GFT_NATURETYPE="' + NatureType + '"', True);
  if not QTyp.EOF then
  begin
    THValComboBox(GetControl('GFM_TYPETARIF')).Value := QTyp.FindField('GFT_CODETYPE').AsString;
    THValComboBox(GetControl('GF_DEVISE')).Value := QTyp.FindField('GFT_DEVISE').AsString;
    THValComboBox(GetControl('GF_DEPOT')).Value := QTyp.FindField('GFT_ETABLISREF').AsString;
  end;
  if QTyp.FindField('GFT_CODETYPE').AsString = '...' then
  begin
    THValComboBox(GetControl('GFM_TYPETARIF')).Value := '...';
    THValComboBox(GetControl('GF_DEVISE')).Value := GetParamSoc('SO_DEVISEPRINC');
    THValComboBox(GetControl('GF_DEPOT')).Value := ''
  end;
  QPer := OpenSQL('Select * from TarifPer where GFP_DERUTILISE="X"', True);
  if not QTyp.EOF then
  begin
    THValComboBox(GetControl('GFM_PERTARIF')).Value := QPer.FindField('GFP_CODEPERIODE').AsString;
    THEdit(GetControl('GF_DATEDEBUT')).Text := QPer.FindField('GFP_DATEDEBUT').AsString;
    THEdit(GetControl('GF_DATEFIN')).Text := QPer.FindField('GFP_DATEFIN').AsString;
  end;
  Ferme(QTyp);
  Ferme(QPer);
end;

// Suppression de tarif par lot

procedure TOF_TARIFSUPPRESSION.SuppressionTarif;
var F: TFMul;
  i, CodeTarif, TarfMode: integer;
  TousCodeTarif, CodeArticle, Depot: string;
begin
  if HShowMessage('0;Confirmation;Confirmez vous l''épuration des tarifs?;Q;YN;N;N;', '', '') <> mrYes then exit;
  F := TFMul(Ecran);
  if (F.FListe.NbSelected = 0) and (not F.FListe.AllSelected) then
  begin
    //MessageAlerte('Aucun élément sélectionné');
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    if VAlerte <> nil then VAlerte.Visible := FALSE;
    {$ENDIF}
    HShowMessage('0;' + F.Caption + ';Aucun élément sélectionné;W;O;O;O;', '', '');
    exit;
  end;
  with TFMul(Ecran) do
  begin
    if FListe.AllSelected then
    begin
      Q.First;
      while not Q.EOF do
      begin
        CodeTarif := TFmul(Ecran).Q.FindField('GF_TARIF').asInteger;
        CodeArticle := TRIM(copy(TFmul(Ecran).Q.FindField('GF_ARTICLE').asString, 1, 18));
        TarfMode := TFmul(Ecran).Q.FindField('GF_TARFMODE').asInteger;
        Depot := TFmul(Ecran).Q.FindField('GF_DEPOT').AsString;
        RechercheToutTarif(TousCodeTarif, CodeArticle, Depot, TarfMode);
        if TousCodeTarif = '' then TousCodeTarif := IntToStr(CodeTarif)
        else TousCodeTarif := TousCodeTarif + ',' + IntToStr(CodeTarif);
        Q.NEXT;
      end;
      FListe.AllSelected := False;
    end else
    begin
      for i := 0 to FListe.NbSelected - 1 do
      begin
        FListe.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row - 1);
        {$ELSE}
        {$ENDIF}
        CodeTarif := TFmul(Ecran).Q.FindField('GF_TARIF').asInteger;
        CodeArticle := TRIM(copy(TFmul(Ecran).Q.FindField('GF_ARTICLE').asString, 1, 18));
        TarfMode := TFmul(Ecran).Q.FindField('GF_TARFMODE').asInteger;
        Depot := TFmul(Ecran).Q.FindField('GF_DEPOT').AsString;
        RechercheToutTarif(TousCodeTarif, CodeArticle, Depot, TarfMode);
        if TousCodeTarif = '' then TousCodeTarif := IntToStr(CodeTarif)
        else TousCodeTarif := TousCodeTarif + ',' + IntToStr(CodeTarif);
      end;
      FListe.ClearSelected;
    end;
  end;
  // Assistant de mise à jour
  SupprimerTousLesTarifs(TousCodeTarif);
  if F.FListe.AllSelected then F.FListe.AllSelected := False else F.FListe.ClearSelected;
  F.bSelectAll.Down := False;
end;

procedure RechercheToutTarif(var TousCodeTarif, CodeArticle, Depot: string; TarfMode: Integer);
var SQLTarif: string;
  QTarif: TQuery;
begin
  SQLTarif := ('SELECT * FROM TARIF WHERE GF_ARTICLE LIKE "' + CodeArticle + '%" AND GF_TARFMODE="' + IntToStr(TarfMode) + '" AND GF_DEPOT="' + Depot + '"');
  QTarif := OpenSQL(SQLTarif, True);
  while not QTarif.EOF do
  begin
    if TousCodeTarif = '' then TousCodeTarif := IntToStr(QTarif.FindField('GF_TARIF').AsInteger)
    else TousCodeTarif := TousCodeTarif + ',' + IntToStr(QTarif.FindField('GF_TARIF').AsInteger);
    QTarif.Next;
  end;
  Ferme(QTarif);
end;

procedure SupprimerTousLesTarifs(CodeTarif: string);
var SQLTarif: string;
begin
  SQLTarif := 'update tarif set GF_FERME="X",GF_DATEMODIF="' + USTime(NowH) + '" where gf_tarif in (' + CodeTarif + ') and gf_ferme="-"';
  ExecuteSQL(SQLTarif);
end;

procedure AGLOnChangeCombo(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFMul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_TARIFSUPPRESSION) then TOF_TARIFSUPPRESSION(TOTOF).OnChangeCombo(Parms[1]);
end;

/////////////// Procedure appellé par le bouton Validation //////////////

procedure AGLSuppressionTarif(parms: array of variant; nb: integer);
var F: TForm;
  MaTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then MaTOF := TFMul(F).LaTOF else exit;
  if (MaTOF is TOF_TARIFSUPPRESSION) then TOF_TARIFSUPPRESSION(MaTOF).SuppressionTarif else exit;
end;

initialization
  registerclasses([TOF_TARIFSUPPRESSION]);
  RegisterAglProc('OnChangeEntete', TRUE, 1, AGLOnChangeCombo);
  RegisterAglProc('SuppressionTarif', TRUE, 0, AGLSuppressionTarif);
end.
