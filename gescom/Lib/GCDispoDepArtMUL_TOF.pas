{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCDispoDepArtMUL ()
Mots clefs ... : TOF;GCDispoDepArtMUL
*****************************************************************}
Unit GCDispoDepArtMUL_TOF ;

Interface

{$IFDEF STK}

Uses
  StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
    db,
    dbtables,
    Fe_Main,
    Mul,
  {$ELSE}
    MainEAgl,
    eMul,
  {$ENDIF}
  Menus,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  WTOF,
  Dispo,
  DispoDetail,
  wMnu
  ;

Type
  TOF_GCDispoDepArtMUL = Class(twTOF)
    procedure OnArgument(S: String); override;
    procedure OnNew                ; override;
    procedure OnLoad               ; override;
    procedure OnUpdate             ; override;
    procedure OnDelete             ; override;
    procedure OnClose              ; override;
  private
    PmFlux: TPopupMenuFlux;
    {$IFDEF EAGLCLIENT}
      sColWidths: String;
    {$ENDIF}

    procedure MNFlux_OnClick(Sender: TObject);
    procedure FListe_OnDblClick(sender: TObject);

    { Get }
    function GetCleGQ: tCleGQ;

    { Set }
    procedure SetColsVisible;

    { Make }
    function MakeWhere: string;

    { Action }
    procedure AcEntreeGQ_OnClick(Sender: TObject);
    procedure AcSortieGQ_OnClick(Sender: TObject);

    { Historique }
    {$IFDEF STK}
    procedure HiGSMPhysique_OnClick(Sender: TObject);
    {$ENDIF}

    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);
    {$IFDEF STK}
    procedure LpPhysiqueGQ_OnClick(Sender: TObject);
    procedure LpReserveGQ_OnClick(Sender: TObject);
    procedure LpAttenduGQ_OnClick(Sender: TObject);
    procedure LpProjeteGQ_OnClick(Sender: TObject);
    {$ENDIF}

    { Utilitaires }
    procedure UtRecalculCompteursGQ_OnClick(Sender: TObject);
  public

  end;

{$ENDIF}

Implementation

{$IFDEF STK}

uses
  UtilArticle,
  wCommuns,
  HTB97,
  {$IFDEF STK}
    StkMouvement,
  {$ENDIF}
  wAction
  ;

procedure TOF_GCDispoDepArtMUL.OnArgument(S: String);
begin
  { pour wTof }
  FTableName := 'DISPO';

  Inherited;

  FNature := 'GC';

  { Unités }
  if Assigned(GetControl('FLISTE')) then
    THGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;

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

  { Utilitaires }
  if Assigned(GetControl('UtRecalculCompteursGQ')) then
    TMenuItem(GetControl('UtRecalculCompteursGQ')).OnClick := UtRecalculCompteursGQ_OnClick;

  { Unité de flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));

  { Sélection }
  wSelectAll := true;
end;

procedure TOF_GCDispoDepArtMUL.OnNew;
begin
  Inherited;
end;

procedure TOF_GCDispoDepArtMUL.OnLoad;
begin
  Inherited;
end;

procedure TOF_GCDispoDepArtMUL.OnUpdate;
{$IFDEF EAGLCLIENT}
var
  i: Integer;
{$ENDIF}
begin
  inherited;

  {$IFDEF EAGLCLIENT}
    sColWidths := '';
    for i := 1 to Pred(TFMul(Ecran).FListe.ColCount) do
      sColWidths := sColWidths + IntToStr(TFMul(Ecran).FListe.ColWidths[i]) + ';';
  {$ENDIF}

  SetColsVisible;
end;

procedure TOF_GCDispoDepArtMUL.OnDelete;
begin
  Inherited;
end;

procedure TOF_GCDispoDepArtMUL.OnClose;
begin
  Inherited;

  PmFlux.Free;
end;

procedure TOF_GCDispoDepArtMUL.MNFlux_OnClick(Sender: TObject);
begin
  PmFlux.Flux := StringReplace(TPopupMenu(Sender).Name, 'MN', '', [rfIgnoreCase]);
  SetColsVisible;
end;

procedure TOF_GCDispoDepArtMUL.SetColsVisible;
var
  i: Integer;
  {$IFDEF EAGLCLIENT}
    sTitresCols, sWidths, s2: String;
    OldWidth: Integer;
  {$ENDIF}

  procedure SetColVisible(ColName: String);
 {$IFDEF EAGLCLIENT}
    var
      Visibility: Boolean;
 {$ENDIF}
  begin
    {$IFDEF EAGLCLIENT}
      if Copy(s2, 1, Length(ColName)) = ColName then
      begin
        Visibility := s2 = ColName + PmFlux.Flux;
        if not Visibility then
          TFMul(Ecran).FListe.ColWidths[i] := -1
        else
          TFMul(Ecran).FListe.ColWidths[i] := OldWidth;
      end;
    {$ELSE}
      if Copy(TFMul(Ecran).FListe.Columns[i].FieldName, 1, Length(ColName)) = ColName then
        TFMul(Ecran).FListe.Columns[i].Visible := TFMul(Ecran).FListe.Columns[i].FieldName = ColName + PmFlux.Flux;
    {$ENDIF}
  end;

begin
  {$IFDEF EAGLCLIENT}
    sTitresCols := TFMul(Ecran).FListe.Titres.GetText;
    sWidths := sColWidths;
    for i:= 1 to TFMul(Ecran).FListe.ColCount - 1 do
  {$ELSE}
    for i := 0 to TFMul(Ecran).FListe.Columns.Count - 1 do
  {$ENDIF}
  begin
    {$IFDEF EAGLCLIENT}
      s2 := ReadTokenSt(sTitresCols);
      OldWidth := ValeurI(ReadTokenSt(sWidths));
    {$ENDIF}

    SetColVisible('PHYSIQUE');
    SetColVisible('RESERVE');
    SetColVisible('DISPOABS');
    SetColVisible('ATTENDU');
    SetColVisible('PROJABS');
    SetColVisible('LIBUNITE');
  end;
end;

procedure TOF_GCDispoDepArtMUL.AcEntreeGQ_OnClick(Sender: TObject);
begin
  EntreeGQD(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;

procedure TOF_GCDispoDepArtMUL.AcSortieGQ_OnClick(Sender: TObject);
begin
  SortieGQD(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;

{$IFDEF STK}
procedure TOF_GCDispoDepArtMUL.LpPhysiqueGQ_OnClick(Sender: TObject);
begin
  CallMulGQD(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;
{$ENDIF}

{$IFDEF STK}
procedure TOF_GCDispoDepArtMUL.LpReserveGQ_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;
begin
  CallResGSM(GetArgument , PmFlux.Flux);
  RefreshDB;
end;
{$ENDIF}

{$IFDEF STK}
procedure TOF_GCDispoDepArtMUL.LpAttenduGQ_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;
begin
  CallAttGSM(GetArgument , PmFlux.Flux);
  RefreshDB;
end;
{$ENDIF}

{$IFDEF STK}
procedure TOF_GCDispoDepArtMUL.LpProjeteGQ_OnClick(Sender: TObject);
  function GetArgument: String;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticlefromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;
begin
  CallProGSM(GetArgument, PmFlux.Flux);
  RefreshDB;
end;
{$ENDIF}
        
{$IFDEF STK}
procedure TOF_GCDispoDepArtMUL.HiGSMPhysique_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + wGetCodeArticleFromArticle(GetString('GQ_ARTICLE'))
            + ';DEPOT=' + GetString('GQ_DEPOT')
  end;
begin
  CallPhyGSM(GetArgument, PmFlux.Flux);
  RefreshDB;
end;
{$ENDIF}

procedure TOF_GCDispoDepArtMUL.FListe_OnDblClick(sender: TObject);
begin
  if not IsEmpty then
  begin
    CallFicGQ(GetCleGQ, PmFlux.Flux, Action, GetString('UNITESTO'), GetString('UNITEVTE'), GetString('UNITEACH'), GetString('UNITEPRO'), GetString('UNITECON'));
    RefreshDB;
  end;
end;

function TOF_GCDispoDepArtMUL.GetCleGQ: TCleGQ;
begin
  Result.Article := GetString('GQ_ARTICLE');
  Result.Depot := GetString('GQ_DEPOT');
  Result.Cloture := GetBoolean('GQ_CLOTURE');
  Result.DateCloture := GetDate('GQ_DATECLOTURE');
end;

procedure TOF_GCDispoDepArtMUL.MnLpArticle_OnClick(Sender: TObject);
  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GQ_ARTICLE');
  end;
begin
	wCallGA(GetRange);
  RefreshDb;
end;

procedure TOF_GCDispoDepArtMUL.UtRecalculCompteursGQ_OnClick(Sender: TObject);

  function GetArgument: string;
  begin
    Result := 'WHERE=(' + MakeWhere + ')';
  end;

begin
  if PgiAsk(TraduireMemoire('Confirmez-vous le recalcul des lignes sélectionnées?'), Ecran.Caption) = MrYes then
    if wDoAction(wtaRecalculCompteursGQ, GetArgument) then
      RefreshDB;
end;

function TOF_GCDispoDepArtMUL.MakeWhere: string;
begin
 {$IFNDEF EAGLCLIENT}
  	Result := wMakeWhereFromList('GQ', TFMul(Ecran).FListe, TFMul(Ecran).Q);
 {$ELSE}
  	Result := wMakeWhereFromList('GQ', TFMul(Ecran).FListe, TFMul(Ecran).Q.TQ);
 {$ENDIF}
end;

Initialization
  registerclasses([TOF_GCDispoDepArtMUL]);

{$ENDIF}  
end.
