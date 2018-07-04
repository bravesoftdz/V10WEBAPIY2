{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GSMPHYSIQUE ()
Mots clefs ... : TOF;GSMPHYSIQUE
*****************************************************************}
Unit GSMPHYSIQUE_TOF ;

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
  HTB97,
  Menus,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  wTof,
  Dispo,
  wMnu,
  stkMouvement
  ;

Type
  TOF_GSMPHYSIQUE = Class (tWTOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    PmFLux: TPopupMenuFlux;

    { Evenements }
    procedure FLISTE_OnDblClick(Sender: TOBject);

    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);
    procedure MnLpDispo_OnClick(Sender: TObject);
    procedure MnLpRefOrigine_OnClick(Sender: TObject);

    { Get }
    function GetCleGQ: tCleGQ;
    function GetCleGSM: tCleGSM;
    function GetTitre: string;

    { Flux }
    procedure MNFlux_OnClick(Sender: TObject);
  end ;

{$ENDIF}

Implementation

{$IFDEF STK}

uses
  wCommuns,
  UtilArticle
  ;

procedure TOF_GSMPHYSIQUE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GSMPHYSIQUE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GSMPHYSIQUE.OnUpdate ;
begin
  Inherited ;
  SetColsVisible;
end ;

procedure TOF_GSMPHYSIQUE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GSMPHYSIQUE.OnArgument (S : String ) ;
begin
  FTableName := 'STKMOUVEMENT';

  Inherited ;

  { titre }
  if GetTitre <> '' then Ecran.Caption := GetTitre;


  { Unité de flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));

  { Gestion des colonnes }
  InitParamsForColumns('PHYSIQUE;LIBUNITE', PmFlux);

  { Evenements }
  if Assigned(GetControl('FLISTE')) then
    THGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;

  { Loupe }
  if Assigned(GetControl('MNLPARTICLE')) then
    TMenuItem(GetControl('MNLPARTICLE')).OnClick := MnLpArticle_OnClick;
  if Assigned(GetControl('MNLPDISPO')) then
    TMenuItem(GetControl('MNLPDISPO')).OnClick := MnLPDispo_OnClick;
  if Assigned(GetControl('MNLPREFORIGINE')) then
    TMenuItem(GetControl('MNLPREFORIGINE')).OnClick := MnLPRefOrigine_OnClick;

  { bAgrandir }
  if Assigned(GetControl('BAGRANDIR')) then
    ttoolbarButton97(GetControl('BAGRANDIR')).Click;
end ;

procedure TOF_GSMPHYSIQUE.OnClose ;
begin
  Inherited ;

  PmFlux.Free;
end ;

procedure TOF_GSMPHYSIQUE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_GSMPHYSIQUE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_GSMPHYSIQUE.MnLpArticle_OnClick(Sender: TObject);
  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GSM_ARTICLE');
  end;
begin
	wCallGA(GetRange);
  RefreshDB;
end;

procedure TOF_GSMPHYSIQUE.MnLpDispo_OnClick(Sender: TObject);
begin
	CallMulGQ(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;

function TOF_GSMPHYSIQUE.GetCleGQ: tCleGQ;
begin
  Result.Article := GetString('GSM_ARTICLE');
  Result.Depot := GetString('GSM_DEPOT');
  Result.Cloture := False;
  Result.DateCloture := iDate1900;
end;

function TOF_GSMPHYSIQUE.GetCleGSM: tCleGSM;
begin
  Result.StkTypeMVT  := GetString('GSM_STKTYPEMVT');
  Result.QualifMVT   := GetString('GSM_QUALIFMVT');
  Result.Identifiant := GetInteger('GSM_IDENTIFIANT');
end;

procedure TOF_GSMPHYSIQUE.MNFlux_OnClick(Sender: TObject);
begin
  PmFlux.Flux := StringReplace(TPopupMenu(Sender).Name, 'MN', '', [rfIgnoreCase]);
  SetColsVisible;
end;

procedure TOF_GSMPHYSIQUE.MnLpRefOrigine_OnClick(Sender: TObject);
begin
  CallFicRefOrigine(GetString('GSM_REFORIGINE'));
  RefreshDB;
end;

procedure TOF_GSMPHYSIQUE.FLISTE_OnDblClick(Sender: TOBject);
begin
  if not IsEmpty then
  begin
    CallFicGSM(GetCleGSM, PmFlux.Flux, Action, GetString('UNITESTO'), GetString('UNITEVTE'), GetString('UNITEACH'), GetString('UNITEPRO'), GetString('UNITECON'));
    RefreshDB;
  end;
end;

function TOF_GSMPHYSIQUE.GetTitre: string;
begin
  Result := GetArgumentValue(StArgument, 'TITRE', false);
end;

Initialization
  registerclasses ( [ TOF_GSMPHYSIQUE ] ) ;

{$ENDIF}
  
end.
