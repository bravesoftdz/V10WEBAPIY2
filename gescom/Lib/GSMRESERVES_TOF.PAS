{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GSMRESERVES ()
Mots clefs ... : TOF;GSMRESERVES
*****************************************************************}
Unit GSMRESERVES_TOF ;

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
  wMnu
  ;

Type
  TOF_GSMRESERVES = Class (tWTOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    PmFlux: TPopupMenuFlux;

    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);
    procedure MnLpDispo_OnClick(Sender: TObject);
    procedure MnLpRefOrigine_OnClick(Sender: TObject);

    { Get }
    function GetCleGQ: tCleGQ;

    { Flux }
    procedure MNFlux_OnClick(Sender: TObject);
  end ;

{$ENDIF}

Implementation

{$IFDEF STK}

uses
  wCommuns,
  UtilArticle,
  StkMouvement
  ;

procedure TOF_GSMRESERVES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GSMRESERVES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GSMRESERVES.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GSMRESERVES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GSMRESERVES.OnArgument (S : String ) ;
begin
  FTableName := 'STKMOUVEMENT';

  Inherited ;

  { Unité de flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));

  { Gestion des colonnes }
  InitParamsForColumns('QPREVUE;PHYSIQUE;QRESTE;LIBUNITE', PmFlux);

  { Loupe }
  if Assigned(GetControl('MNLPARTICLE')) then TMenuItem(GetControl('MNLPARTICLE')).OnClick := MnLpArticle_OnClick;
  if Assigned(GetControl('MNLPDISPO')) then TMenuItem(GetControl('MNLPDISPO')).OnClick := MnLPDispo_OnClick;
  if Assigned(GetControl('MNLPREFORIGINE')) then TMenuItem(GetControl('MNLPREFORIGINE')).OnClick := MnLPRefOrigine_OnClick;

  { bAgrandir }
  if Assigned(GetControl('BAGRANDIR')) then ttoolbarButton97(GetControl('BAGRANDIR')).Click;
end ;

procedure TOF_GSMRESERVES.OnClose ;
begin
  Inherited ;

  PmFlux.Free;
end ;

procedure TOF_GSMRESERVES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_GSMRESERVES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_GSMRESERVES.MnLpArticle_OnClick(Sender: TObject);
  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GSM_ARTICLE');
  end;
begin
	wCallGA(GetRange);
  RefreshDB;
end;

procedure TOF_GSMRESERVES.MnLpDispo_OnClick(Sender: TObject);
begin
	CallMulGQ(GetCleGQ, PmFlux.Flux);
  RefreshDB;
end;

function TOF_GSMRESERVES.GetCleGQ: tCleGQ;
begin
  Result.Article := GetString('GSM_ARTICLE');
  Result.Depot := GetString('GSM_DEPOT');
  Result.Cloture := False;
  Result.DateCloture := iDate1900;
end;

procedure TOF_GSMRESERVES.MNFlux_OnClick(Sender: TObject);
begin
  PmFlux.Flux := StringReplace(TPopupMenu(Sender).Name, 'MN', '', [rfIgnoreCase]);
  SetColsVisible;
end;

procedure TOF_GSMRESERVES.MnLpRefOrigine_OnClick(Sender: TObject);
begin
  CallFicRefOrigine(GetString('GSM_REFORIGINE'));
  RefreshDB;
end;

Initialization
  registerclasses ( [ TOF_GSMRESERVES ] ) ;

{$ENDIF}  
end.
