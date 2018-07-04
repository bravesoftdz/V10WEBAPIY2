{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DISPODETAIL (DISPODETAIL)
Mots clefs ... : TOM;DISPODETAIL
*****************************************************************}
Unit DISPODETAIL_TOM ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  forms,
  sysutils,
  ComCtrls,
  Dialogs,
  Menus,
  hDb,

  {$IFNDEF EAGLCLIENT}
    db,
    dbTables,
    Fe_Main,
    Fiche,
    Mul,
  {$ELSE}
    MainEagl,
    eFiche,
    eMul,
  {$ENDIF}

  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  HTB97,
  wTom,
  EntGP,
  wCommuns,
  wMnu
  ;

Type
  TOM_DISPODETAIL = Class (tWTOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    PmFLux: TPopupMenuFlux;

    { événements }
    procedure MNFlux_OnClick(Sender: TObject);

    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);

    { Historique }
    {$IFDEF STK}
      procedure HiGSMPhysique_OnClick(Sender: TObject);
    {$ENDIF}

    { Set }
    procedure SetCodeArticle;

    { Get }
    function GetUniteFromFlux: String;

  end ;

Implementation

uses
  UtilArticle,
  stkMouvement
  ;

procedure TOM_DISPODETAIL.OnNewRecord ;
begin
  Inherited ;
  { Statut de disponibilité: Libre par défaut }
  SetString('GQD_STATUTDISPO', 'LBR');

  { Statut de flux: Standard par défaut }
  SetString('GQD_STATUTFLUX', 'STD');

  { Date entrée du lot: Date système (écrasée su nécéssaire) }
  SetDate('GQD_DATEENTREELOT', V_pgi.DateEntree);
end ;

procedure TOM_DISPODETAIL.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_DISPODETAIL.OnUpdateRecord ;
begin
  Inherited ;

	DisableControl;
  try
    SetCodeArticle;
  finally
    EnableControl;
  end;
end ;

procedure TOM_DISPODETAIL.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_DISPODETAIL.OnLoadRecord ;
begin
  Inherited ;

  SetControlText('UNITEFLUX', GetUniteFromFlux);

  Ecran.Caption := TraduireMemoire('Mouvement de stock de l''article ') + ': ' + wGetCodeArticleFromArticle(GetString('GQD_CODEARTICLE') + ', ' + TraduireMemoire('dépôt') + ': ' + RechDom('GCDEPOT', GetString('GQD_DEPOT'), False));
end ;

procedure TOM_DISPODETAIL.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_DISPODETAIL.OnArgument ( S: String ) ;
begin
  Inherited ;

  { Flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));

  {$IFDEF STK}
  { Historique }
  if Assigned(GetControl('HiGSMPhysique')) then
    TMenuItem(GetControl('HiGSMPhysique')).OnClick := HiGSMPhysique_OnClick;
  {$ENDIF}

  { Loupe }
  if Assigned(GetControl('MNLPARTICLE')) then TMenuItem(GetControl('MNLPARTICLE')).OnClick := MnLpArticle_OnClick;
end ;

procedure TOM_DISPODETAIL.OnClose ;
begin
  Inherited ;
  PmFlux.Free;
end ;

procedure TOM_DISPODETAIL.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_DISPODETAIL.SetCodeArticle;
begin
	SetString('GQD_CODEARTICLE', wGetCodeArticleFromArticle(GetString('GQD_ARTICLE')));
end;

procedure TOM_DISPODETAIL.MNFlux_OnClick(Sender: TObject);
var
  i: Integer;
begin
  PmFLux.Flux := StringReplace(TPopupMenu(Sender).Name, 'MN', '', [rfIgnoreCase]);

  for i := 0 to DS.FieldCount - 1 do
    CalculField(DS.Fields[i].FieldName);

  SetControlText('UNITEFLUX', GetUniteFromFlux);
end;


procedure TOM_DISPODETAIL.MnLpArticle_OnClick(Sender: TObject);
  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GQD_ARTICLE');
  end;
begin
	wCallGA(GetRange);
  RefreshDB;
end;

function TOM_DISPODETAIL.GetUniteFromFlux: String;
begin
  Result := GetArgumentValue(stArgument, 'UNITE' + PmFLux.Flux);
end;

{$IFDEF STK}
procedure TOM_DISPODETAIL.HiGSMPhysique_OnClick(Sender: TObject);
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

Initialization
  registerclasses ( [ TOM_DISPODETAIL ] ) ;
end.
