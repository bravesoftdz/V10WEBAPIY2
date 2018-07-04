{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 14/05/2003
Description .. : TOM de la TABLE : DISPO
*****************************************************************}
Unit DISPO_TOM;

Interface

{$IFDEF STK}

Uses
  StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
    db,
    dbTables,
    Fe_Main,
    Fiche,
    FichList,
  {$ELSE}
    MainEagl,
    eFiche,
    eFichList,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  WTOM,
  Dispo,
  DispoDetail,
  wMnu
  ;

Type
  TOM_DISPO = Class(twTOM)
    procedure OnArgument(S: String)   ; override;
    procedure OnNewRecord             ; override;
    procedure OnLoadRecord            ; override;
    procedure OnUpdateRecord          ; override;
    procedure OnAfterUpdateRecord     ; override;
    procedure OnCancelRecord          ; override;
    procedure OnDeleteRecord          ; override;
    procedure OnChangeField(F: TField); override;
    procedure OnClose                 ; override;
  private
    PmFlux                              : TPopupMenuFlux;
    UniteSto, FRetour                   : String;
    Loading, FromOnChange, FromCalculQte: Boolean;

    { Get }
    function GetUniteFromFlux: String;
    function GetCleGQ: tCleGQ;

    { Divers }
    procedure CalculQte(FieldName, Unite: String);
    function FormatStringForTHNumEdit(s: String): String;

    { événements }
    procedure MNFlux_OnClick(Sender: TObject);
    procedure StockMin_OnChange(Sender: TObject);
    procedure StockMax_OnChange(Sender: TObject);

    { Action }
    procedure AcEntreeGQ_OnClick(Sender: TObject);
    procedure AcSortieGQ_OnClick(Sender: TObject);

    { Historique }
    {$IFDEF STK}
    procedure HiGSMPhysique_OnClick(Sender: TObject);
    {$ENDIF}

    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);
    procedure LpPhysiqueGQ_OnClick(Sender: TObject);
    procedure LpReserveGQ_OnClick(Sender: TObject);
    procedure LpAttenduGQ_OnClick(Sender: TObject);
    procedure LpProjeteGQ_OnClick(Sender: TObject);
    procedure Loupe_OnClick(Sender: TObject);
    procedure LoupeLBN_OnClick(Sender: TObject);

    { Utilitaires }
    procedure UtRecalculCompteursGQ_OnClick(Sender: TObject);
  protected
    { Control Field }
    procedure ControlField(FieldName: string); override;
    procedure CalculField(FieldName: string); override;
    function  RecordIsValid: boolean; override;
  public

  end;

Const
  TexteMessage : Array[1..1] of String =
               (
                {1} ''
               );

{$ENDIF}

Implementation

{$IFDEF STK}

uses
  wCommuns,
  UtilArticle,
  wConversionUnite,
  EntGP,
  {$IFDEF STK}
    StkMouvement,
  {$ENDIF}
  Menus,
  wAction,
  HTB97,
  extctrls,
  LookUp,
  stkNature
  ;

procedure TOM_DISPO.OnArgument(S: String);
begin
  FTableName := 'DISPO';

  Inherited;

  FromOnChange := False;
  FromCalculQte := False;
  Loading := True;

  UniteSto := GetArgumentValue(StArgument, 'UNITESTO');

  { Stock }
  if Assigned(GetControl('STOCKMIN')) then
    THNumEdit(GetControl('STOCKMIN')).OnChange := StockMin_OnChange;
  if Assigned(GetControl('STOCKMAX')) then
    THNumEdit(GetControl('STOCKMAX')).OnChange := StockMax_OnChange;

  { Action }
  if Assigned(GetControl('AcEntreeGQ')) then
    TMenuItem(GetControl('AcEntreeGQ')).OnClick := AcEntreeGQ_OnClick;
	if Assigned(GetControl('AcSortieGQ')) then
    TMenuItem(GetControl('AcSortieGQ')).OnClick := AcSortieGQ_OnClick;

  { Historique }
  {$IFDEF STK}
  if Assigned(GetControl('HiGSMPhysique')) then
    TMenuItem(GetControl('HiGSMPhysique')).OnClick := HiGSMPhysique_OnClick;
  {$ENDIF}

  { Loupe }
  if Assigned(GetControl('MNLPARTICLE')) then
    TMenuItem(GetControl('MNLPARTICLE')).OnClick := MnLpArticle_OnClick;
  if Assigned(GetControl('LpPhysiqueGQ')) then
    TMenuItem(GetControl('LpPhysiqueGQ')).OnClick := LpPhysiqueGQ_OnClick;
	if Assigned(GetControl('LpReserveGQ')) then
    TMenuItem(GetControl('LpReserveGQ')).OnClick := LpReserveGQ_OnClick;
	if Assigned(GetControl('LpAttenduGQ')) then
    TMenuItem(GetControl('LpAttenduGQ')).OnClick := LpAttenduGQ_OnClick;
	if Assigned(GetControl('LpProjeteGQ')) then
    TMenuItem(GetControl('LpProjeteGQ')).OnClick := LpProjeteGQ_OnClick;

  { Compteurs }
	if Assigned(GetControl('LoupePHY')) then
    TToolBarButton97(GetControl('LoupePHY')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeRES')) then
    TToolBarButton97(GetControl('LoupeRES')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeATT')) then
    TToolBarButton97(GetControl('LoupeATT')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupePRO')) then
    TToolBarButton97(GetControl('LoupePRO')).OnClick := Loupe_OnClick;

  { Entrées }
	if Assigned(GetControl('LoupeEAC')) then
    TToolBarButton97(GetControl('LoupeEAC')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeEPR')) then
    TToolBarButton97(GetControl('LoupeEPR')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeETD')) then
    TToolBarButton97(GetControl('LoupeETD')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeEMA')) then
    TToolBarButton97(GetControl('LoupeEMA')).OnClick := Loupe_OnClick;

  { Sorties }
	if Assigned(GetControl('LoupeSVE')) then
    TToolBarButton97(GetControl('LoupeSVE')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeSPR')) then
    TToolBarButton97(GetControl('LoupeSPR')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeSTD')) then
    TToolBarButton97(GetControl('LoupeSTD')).OnClick := Loupe_OnClick;
	if Assigned(GetControl('LoupeSMA')) then
    TToolBarButton97(GetControl('LoupeSMA')).OnClick := Loupe_OnClick;

  { Libre }
	if Assigned(GetControl('LoupeLB1')) then
    TToolBarButton97(GetControl('LoupeLB1')).OnClick := LoupeLBN_OnClick;
	if Assigned(GetControl('LoupeLB2')) then
    TToolBarButton97(GetControl('LoupeLB2')).OnClick := LoupeLBN_OnClick;
	if Assigned(GetControl('LoupeLB3')) then
    TToolBarButton97(GetControl('LoupeLB3')).OnClick := LoupeLBN_OnClick;
	if Assigned(GetControl('LoupeLB4')) then
    TToolBarButton97(GetControl('LoupeLB4')).OnClick := LoupeLBN_OnClick;
  if Assigned(GetControl('LoupeLB5')) then
    TToolBarButton97(GetControl('LoupeLB4')).OnClick := LoupeLBN_OnClick;

  { Utilitaires }
  if Assigned(GetControl('UtRecalculCompteursGQ')) then
    TMenuItem(GetControl('UtRecalculCompteursGQ')).OnClick := UtRecalculCompteursGQ_OnClick;

  { Flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));
end;

procedure TOM_DISPO.OnNewRecord;
begin
  Inherited;
end;

procedure TOM_DISPO.OnLoadRecord;
begin
  Inherited;
  SetControlText('UNITEFLUX', GetUniteFromFlux);

  Loading := False;
end;

procedure TOM_DISPO.OnUpdateRecord;
begin
  if RecordIsValid then
  begin
    if Assigned(Ecran) then
    begin

    end;
  end;

  Inherited;

  if Assigned(Ecran) then
  begin
    TFFiche(Ecran).Retour := TFFiche(Ecran).Retour + ';' + FRetour;
  end;  
end;

procedure TOM_DISPO.OnAfterUpdateRecord;
begin
  Inherited;
end;

procedure TOM_DISPO.OnCancelRecord;
begin
  Inherited;
end;

procedure TOM_DISPO.OnDeleteRecord;
begin
  Inherited;
end;

procedure TOM_DISPO.OnChangeField(F: TField);
begin
  Inherited;
end;

procedure TOM_DISPO.OnClose;
begin
  Inherited;

  PmFlux.Free;
end;

procedure TOM_DISPO.CalculField(FieldName: string);
begin
  if not FromOnChange then
  begin
    CalculQte(FieldName, GetUniteFromFlux);

    if Pos(FieldName, 'GQ_PHYSIQUE;GQ_RESERVECLI') <> 0 then
      SetControlText('DISPOABSOLU', FormatStringForTHNumEdit(FloatToStr(wConversionSimple(GetDouble('GQ_PHYSIQUE') - GetDouble('GQ_RESERVECLI'), UniteSto, GetUniteFromFlux))))
    else if FieldName = 'GQ_RESERVEFOU' then
      SetControlText('PROJETEABSOLU', FormatStringForTHNumEdit(FloatToStr(wConversionSimple(GetDouble('GQ_PHYSIQUE') - GetDouble('GQ_RESERVECLI') + GetDouble(FieldName), UniteSto, GetUniteFromFlux))));
  end;
end;

procedure TOM_DISPO.ControlField(FieldName: string);
begin
  DisableControl;

  { Type création }
  if Action = 'CREATION' then
  begin

  end;

  EnableControl;

  inherited;
  if LastError <> 0 then
    LastErrorMsg := TexteMessage[LastError];
end;

function TOM_DISPO.RecordIsValid: boolean;
begin
  Result := False;

  { Contrôle bas niveau }
  if not inherited RecordIsvalid then exit;

  { Contrôle }

  Result := (LastError = 0);
  if LastError > 0 then
  begin
    if Assigned(Ecran) then
    begin
      LastErrorMsg := TexteMessage[LastError];
    end
    else
      fTob.AddChampSupValeur('Error', LastErrorMsg, false);
  end;
end;

function TOM_DISPO.GetUniteFromFlux: String;
begin
  Result := GetArgumentValue(stArgument, 'UNITE' + PmFLux.Flux);
end;

procedure TOM_DISPO.CalculQte(FieldName, Unite: String);
var
  FieldSuffixe: String;
begin
  FromCalculQte := True;

  FieldSuffixe := ExtractSuffixe(FieldName);
  if Assigned(GetControl(FieldSuffixe)) and (GetControl(FieldSuffixe) is THNumEdit) then
    SetControlText(FieldSuffixe, FormatStringForTHNumEdit(FloatToStr(wConversionSimple(GetDouble(FieldName), UniteSto, Unite))));

  FromCalculQte := False;
end;

function TOM_DISPO.FormatStringForTHNumEdit(s: String): String;
var
  sAfterComa: String;
begin
  if Pos(',', s) = 0 then
    Result := s + ',000'
  else
  begin
    sAfterComa := Copy(s, Pos(',', s) + 1, Length(s));
    Result := s + wStringRepeat('0', 3 - Length(sAfterComa));
  end;
end;

procedure TOM_DISPO.StockMin_OnChange(Sender: TObject);
begin
  if not FromCalculQte then
  begin
    if DS.State = dsBrowse then
      DS.Edit;

    FromOnChange := True;
    SetDouble('GQ_STOCKMIN', wConversionSimple(Valeur(GetControlText('STOCKMIN')), GetUniteFromFlux, UniteSto));
    FromOnChange := False;
  end;
end;

procedure TOM_DISPO.StockMax_OnChange(Sender: TObject);
begin
  if not FromCalculQte then
  begin
    if DS.State = dsBrowse then
      DS.Edit;

    FromOnChange := True;
    SetDouble('GQ_STOCKMAX', wConversionSimple(Valeur(GetControlText('STOCKMAX')), GetUniteFromFlux, UniteSto));
    FromOnChange := False;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 16/05/2003
Modifié le ... :   /  /
Description .. : Appel de la fiche article
Mots clefs ... :
*****************************************************************}
procedure TOM_DISPO.MnLpArticle_OnClick(Sender: TObject);

  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GQ_ARTICLE');
  end;

begin
	wCallGA(GetRange);
  RefreshDB;
end;

procedure TOM_DISPO.AcEntreeGQ_OnClick(Sender: TObject);
begin
  EntreeGQD(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;

procedure TOM_DISPO.AcSortieGQ_OnClick(Sender: TObject);
begin
  RefreshDB;
end;

procedure TOM_DISPO.LpPhysiqueGQ_OnClick(Sender: TObject);
begin
  CallMulGQD(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;

procedure TOM_DISPO.LpReserveGQ_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;
begin
  CallResGSM(GetArgument , PmFLux.Flux);
  RefreshDB;
end;

procedure TOM_DISPO.LpAttenduGQ_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;
begin
  CallAttGSM(GetArgument, PmFLux.Flux);
  RefreshDB;
end;

procedure TOM_DISPO.LpProjeteGQ_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;
begin
  CallProGSM(GetArgument, PmFLux.Flux);
  RefreshDB;
end;


{$IFDEF STK}
procedure TOM_DISPO.HiGSMPhysique_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + GetString('GQD_CODEARTICLE')
            + ';DEPOT=' + GetString('GQD_DEPOT')
  end;
begin
  CallPhyGSM(GetArgument, PmFLux.Flux);
  RefreshDB;
end;
{$ENDIF}

procedure TOM_DISPO.MNFlux_OnClick(Sender: TObject);
var
  i: Integer;
begin
  PmFLux.Flux := StringReplace(TPopupMenu(Sender).Name, 'MN', '', [rfIgnoreCase]);

  for i := 0 to DS.FieldCount - 1 do
    CalculField(DS.Fields[i].FieldName);

  SetControlText('UNITEFLUX', GetUniteFromFlux);
end;

function TOM_DISPO.GetCleGQ: tCleGQ;
begin
  Result.Article := GetString('GQ_ARTICLE');
  Result.Depot := GetString('GQ_DEPOT');
  Result.Cloture := GetBoolean('GQ_CLOTURE');
  Result.DateCloture := GetDate('GQ_DATECLOTURE');
end;

procedure TOM_DISPO.UtRecalculCompteursGQ_OnClick(Sender: TObject);

  function GetArgument: string;
  begin
    Result := 'WHERE=' + WhereGQ(GetCleGQ);
  end;

begin
  wDoAction(wtaRecalculCompteursGQ, GetArgument);
  RefreshDB;
end;

function CleGQToStr(CleGQ: TCleGQ): String;
begin
  Result := CleGQ.Article + '~' + CleGQ.Depot + '~' + BoolToStr(CleGQ.Cloture) + '~' + DateToStr(CleGQ.DateCloture);
end;

procedure TOM_DISPO.Loupe_OnClick(Sender: TObject);
var
  QualifMvt: string;

  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')

  end;
begin
  QualifMvt := wRight(TToolBarButton97(Sender).Name, 3);

  if      QualifMvt = 'PHY' then CallMulGQD(GetCleGQ, PmFlux.Flux)
  else if QualifMvt = 'RES' then CallResGSM(GetArgument, PmFLux.Flux)
  else if QualifMvt = 'ATT' then CallAttGSM(GetArgument, PmFLux.Flux)
  else if QualifMvt = 'PRO' then CallProGSM(GetArgument, PmFlux.Flux)
  else if Pos(QualifMvt, 'EAC;EPR;ETD;EMA;SVE;SPR;STD;SMA') > 0 then CallPhyGSM(GetArgument + ';QUALIFMVT=' + wRight(TToolBarButton97(Sender).Name, 3), PmFLux.Flux)
  else exit;

  RefreshDB;
end;

procedure TOM_DISPO.LoupeLBN_OnClick(Sender: TObject);
var
  iGSN                            : integer;
  StkTypeMvt, Compteur, QualifsMvt: string;

  function GetArgument: string;
  begin
    Result := 'QUALIFMVT=' + QualifsMvt
            + ';CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;

begin
  Compteur := Copy(TToolBarButton97(Sender).Name, 6, 3);

  { Déduction des natures de mouvements }
  GetTobGSN;
  for iGSN := 0 to Vh_Gp.TobGSN.Detail.Count-1 do
  begin
    if Pos(Compteur, Vh_Gp.TobGSN.Detail[iGSN].G('GSN_QTEPLUS')) > 0 then
    begin
      QualifsMvt := QualifsMvt + iif(QualifsMvt <> '', '~', '') + Vh_Gp.TobGSN.Detail[iGSN].G('GSN_QUALIFMVT');
      StkTypeMvt := Vh_Gp.TobGSN.Detail[iGSN].G('GSN_STKTYPEMVT')
    end;
  end;

  { Appel du bon mul }
  if      StkTypeMvt = 'PHY' then CallPhyGSM(GetArgument, PmFLux.Flux)
  else if StkTypeMvt = 'RES' then CallResGSM(GetArgument, PmFLux.Flux)
  else if StkTypeMvt = 'ATT' then CallAttGSM(GetArgument, PmFLux.Flux)
end;

Initialization
  registerclasses([TOM_DISPO]);

{$ENDIF}
end.
