{***********UNITE*************************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 20/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCDispoDetail ()
Mots clefs ... : TOF;GCDispoDetail
*****************************************************************}
Unit GCDispoDetail_TOF ;

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
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  WTOF,
  wMnu,
  uTob,
  DispoDetail
  ;

Type
  MyArrayInt = Array of Integer;

  TOF_GCDispoDetail = Class(twTOF)
    procedure OnArgument(S: String); override;
    procedure OnNew                ; override;
    procedure OnLoad               ; override;
    procedure OnUpdate             ; override;
    procedure OnDelete             ; override;
    procedure OnClose              ; override;
  private
    PmFlux: TPopupMenuFlux;

    { Evenements }
    procedure FLISTE_OnDblClick(Sender: TOBject);

    { Flux }
    procedure MNFlux_OnClick(Sender: TObject);


    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);

    { Historique }
    {$IFDEF STK}
    procedure HiGSMPhysique_OnClick(Sender: TObject);
    {$ENDIF}
  public

  end;

{$ENDIF}

Implementation

{$IFDEF STK}

uses
  wCommuns,
  Menus,
  UtilArticle,
  wAction,
  StkMouvement
  ;

procedure TOF_GCDispoDetail.OnArgument(S: String);
begin
  { wTof }
  FNature    := 'GC';
  FTableName := 'DISPODETAIL';

  Inherited;

  { Evenements }
  if Assigned(GetControl('FLISTE')) then THGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;

  { Unité de flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));

  { Gestion des colonnes }
  InitParamsForColumns('PHYSIQUE;LIBUNITE', PmFlux, 'GQD_REFAFFECTATION;GQD_EMPLACEMENT;GQD_LOTINTERNE;GQD_LOTEXTERNE;GQD_SERIEINTERNE;GQD_SERIEEXTERNE;'
                                                    + 'GQD_DATEENTREELOT;GQD_DATEDISPO;GQD_DATEPEREMPTION;GQD_INDICEARTICLE;GQD_MARQUE;GQD_REFPROPRIO;'
                                                    + 'GQD_CHOIXQUALITE;');

  { Evénements }
  { Loupe }
  if Assigned(GetControl('MnLpArticle')) then
  	TMenuItem(GetControl('MnLpArticle')).OnClick := MnLpArticle_OnClick;
  { Historique }
  if Assigned(GetControl('HiGSMPhysique')) then
    TMenuItem(GetControl('HiGSMPhysique')).OnClick := HiGSMPhysique_OnClick;

  THEdit(GetControl('GQD_CODEARTICLE')).UpdateLibelle;
  THEdit(GetControl('GQD_EMPLACEMENT')).UpdateLibelle;
end;

procedure TOF_GCDispoDetail.OnNew;
begin
  Inherited;
end;

procedure TOF_GCDispoDetail.OnLoad;
begin
  Inherited;
end;

procedure TOF_GCDispoDetail.OnUpdate;
begin
  inherited;
end;

procedure TOF_GCDispoDetail.OnDelete;
begin
  Inherited;
end;

procedure TOF_GCDispoDetail.OnClose;
begin
  Inherited;

  PmFlux.Free;
end;

procedure TOF_GCDispoDetail.MNFlux_OnClick(Sender: TObject);
begin
  PmFlux.Flux := PmFLux.GetFluxFromSender(Sender);
  SetColsVisible;
end;

procedure TOF_GCDispoDetail.MnLpArticle_OnClick(Sender: TObject);

  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GQD_ARTICLE');
  end;

begin
	wCallGA(GetRange);
  RefreshDB;
end;

{$IFDEF STK}
procedure TOF_GCDispoDetail.HiGSMPhysique_OnClick(Sender: TObject);
  function GetArgument: string;
  begin
    Result := 'CODEARTICLE=' + GetString('GQD_CODEARTICLE')
            + ';DEPOT=' + GetString('GQD_DEPOT')
  end;
begin
  CallPhyGSM(GetArgument, PmFlux.Flux);
  RefreshDB;
end;
{$ENDIF}

procedure TOF_GCDispoDetail.FLISTE_OnDblClick(Sender: TOBject);
begin
  if not IsEmpty then
  begin
    CallFicGQD(GetString('GQD_ARTICLE'),GetString('GQD_DEPOT'),GetInteger('GQD_IDENTIFIANT'), PmFlux.Flux, Action, GetString('UNITESTO'), GetString('UNITEVTE'), GetString('UNITEACH'), GetString('UNITEPRO'), GetString('UNITECON'));
    RefreshDB;
  end;
end;


Initialization
  registerclasses([TOF_GCDispoDetail]);

{$ENDIF}
end.
