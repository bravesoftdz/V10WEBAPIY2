{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : STKMOUVEMENT (STKMOUVEMENT)
Mots clefs ... : TOM;STKMOUVEMENT
*****************************************************************}
Unit STKMOUVEMENT_TOM ;

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
  Math,
  Dispo,
  DispoDetail,
  wMnu
  ;


Type
  tcaract_Art = record
    Gereparlot : Boolean;
    Serialise  : Boolean;
    QualifUniteSto, UniteQteVte,
    UniteQteAch, UniteProd,
    UniteConso : String;

  end;

  TOM_STKMOUVEMENT = Class (tWTOM)
  {$IFDEF STK}
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
    caract_Art : tCaract_Art;
    sCodeArticle, sDepot, sOrig : String;

    { événements }
    procedure MNFlux_OnClick(Sender: TObject);

    { Loupe }
    procedure MnLpArticle_OnClick(Sender: TObject);
    procedure MnLpDispo_OnClick(Sender: TObject);
    procedure MnLpRefOrigine_OnClick(Sender: TObject);

    { Get }
    function GetStkPrevu(AvantModif: boolean): Double;
    function GetPhysique(AvantModif: boolean): Double;
    function GetCleGQD(AvantModif: boolean): tCleGQD;
    function GetCleGQ(AvantModif: boolean): tCleGQ;
    function GetUniteFromFlux: String;

    { Set }
    procedure SetStkTypeMvt;
    procedure SetEtatMvt;
    procedure SetArticle;
    procedure SetCodeArticle;
    procedure SetSigne;
    procedure SetStkConvUnit; { GPAO STKMOUV}
    procedure ChangeFlux(PmFlux : TPopupMenuFlux);

    { Gestion GQD }
    procedure AddToGQD;

    { Gestion GQ }
    procedure AddToGQ;
    procedure RemoveFromGQ(AvantModif: boolean);


    Function getCaractArt(sCodeArticle : string) : tcaract_Art;

  protected
    { Control Field }
    procedure ControlField(FieldName: string); override;
    procedure CalculField(FieldName: string); override;
    function  RecordIsValid: boolean; override;
  {$ENDIF}
  end ;

Const

	TexteMessage: array[1..1] of string = (
	  {1}   'La nature de mouvement n''est pas renseignée'
		);

	ErrQualifMvt = 1;

Implementation

uses
  UtilArticle,
  StkNature,
  StkMouvement
  ;

{$IFDEF STK}

procedure TOM_STKMOUVEMENT.OnNewRecord ;
begin
  Inherited ;
  { Date mouvement }
  SetDate('GSM_DATEMVT', V_Pgi.DateEntree);

  { Statuts }
  SetString('GSM_STATUTDISPO', 'LBR');
  SetString('GSM_STATUTFLUX', 'STD');

  if Assigned(Ecran) then
  begin
    if Ecran.name = 'GSMPHYSIQUE_FIC' then
    begin
      SetString('GSM_CODEARTICLE', sCodeArticle);
      SetString('GSM_DEPOT', sDepot);
      SetString('GSM_QUALIFMVT', 'EMA');
      SetControlProperty('GSM_QUALIFMVT', 'Enabled', False);
    end;
  end;
end ;

procedure TOM_STKMOUVEMENT.OnDeleteRecord ;
begin
  Inherited ;
  { On retire de GQ }
  RemoveFromGQ(false);

  { JL§: Il n'y a pas de RemoveFRomGQD car on ne delete jamais un mouvement physique }
end ;

procedure TOM_STKMOUVEMENT.OnUpdateRecord ;
begin
  Inherited ;

	DisableControl;
  try
    SetArticle;
    SetCodeArticle;
    SetStkConvUnit; { GPAO STKMOUV }
    SetStkTypeMvt;
    SetEtatMvt;
    SetSigne;
  finally
    EnableControl;
  end;
end ;

procedure TOM_STKMOUVEMENT.OnAfterUpdateRecord ;
begin
  Inherited ;
  if Ikc = 'M' then
  begin
    { On retire de GQ }
    RemoveFromGQ(true);

    { JL§: Il n'y a pas de RemoveFromGQD car on ne modifie jamais un mouvement physique }
  end;

  { On cumle dans GQD }
  AddToGQD;

  { On cumule dans les compteurs de GQ }
  AddToGQ;

  { On boucle en création }
  if Assigned(Ecran) then
  begin
    if sOrig = 'MNU' then  TToolBarButton97(GetControl('BINSERT')).click ;
  end;
end ;

procedure TOM_STKMOUVEMENT.OnLoadRecord ;
begin
  Inherited ;
  SetControlText('UNITEFLUX', GetUniteFromFlux);
end ;

Function TOM_STKMOUVEMENT.getCaractArt(sCodeArticle : string) : tcaract_Art;
var acaract_Art : MyArrayValue;
begin
  if sCodeArticle <> '' then
  begin
    if ExisteSQL('SELECT GA_ARTICLE from ARTICLE WHERE GA_CODEARTICLE="'+sCodeArticle+'"') then
    begin
      acaract_Art := WGetSqlFieldsValues(['GA_LOT','GA_NUMEROSERIE','GA_QUALIFUNITESTO','GA_UNITEQTEVTE','GA_UNITEQTEACH','GA_UNITEPROD','GA_UNITECONSO'], 'ARTICLE', 'GA_CODEARTICLE = "'+sCodeArticle+'"');
      Result.Gereparlot := (String(acaract_Art[0])=Wtrue);
      Result.Serialise := (String(acaract_Art[1])=Wtrue);
      Result.QualifUniteSto := (String(aCaract_Art[2]));
      Result.UniteQteVte := (String(acaract_Art[3]));
      Result.UniteQteAch := (String(acaract_Art[4]));
      Result.UniteProd := (String(acaract_Art[5]));
      Result.UniteConso := (String(acaract_Art[6]));
    end;
  end
  else
  begin
    Result.Gereparlot := (true);
    Result.Serialise := (true);
    Result.QualifUniteSto := '';
    Result.UniteQteVte := '';
    Result.UniteQteAch := '';
    Result.UniteProd := '';
    Result.UniteConso := '';
  end;
end;

procedure TOM_STKMOUVEMENT.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  if F.FieldName = 'GSM_CODEARTICLE' then
  begin
    Caract_Art := getCaractArt(GetString('GSM_CODEARTICLE'));
    SetControlProperty('GSM_LOTINTERNE','VISIBLE',caract_Art.Gereparlot);
    SetControlProperty('GSM_LOTEXTERNE','VISIBLE',caract_Art.Gereparlot);
    SetControlProperty('GSM_SERIEINTERNE','VISIBLE',caract_Art.Serialise );
    SetControlProperty('GSM_SERIEEXTERNE','VISIBLE',caract_Art.Serialise );
    SetControlProperty('TGSM_LOTINTERNE','VISIBLE',caract_Art.Gereparlot);
    SetControlProperty('TGSM_LOTEXTERNE','VISIBLE',caract_Art.Gereparlot);
    SetControlProperty('TGSM_SERIEINTERNE','VISIBLE',caract_Art.Serialise );
    SetControlProperty('TGSM_SERIEEXTERNE','VISIBLE',caract_Art.Serialise );
    ChangeFlux(PmFlux);
  end;
end ;

procedure TOM_STKMOUVEMENT.OnArgument ( S: String ) ;
begin
  Inherited ;
  { Flux }
  PmFlux := TPopupMenuFlux.Create(Ecran, MNFlux_OnClick, S, GetArgumentValue(S, 'FLUX'));

  { Gestion des colonnes }
//  ParamsForColumns.FPmFlux := PmFlux;
//  ParamsForColumns.sFieldsList := 'PHYSIQUE;LIBUNITE';

  { Argument }
  sCodeArticle := wgetCodeArticleFromArticle(GetArgumentValue(s, 'ARTICLE'));
  sDepot := GetArgumentValue(s,'DEPOT');
  sOrig := GetArgumentValue(s,'ORIG');

  if GetArgumentvalue(S,'ACTION') = 'CREATION' then
  begin
    SetControlProperty('bInsert','visible',True);
    SetControlProperty('bInsert','actif',True);
  end
  else
  begin
    SetControlProperty('bInsert','visible',False);
    SetControlProperty('bInsert','actif',False);
  end;

  { Loupe }
  if Assigned(GetControl('MNLPARTICLE')) then
    TMenuItem(GetControl('MNLPARTICLE')).OnClick := MnLpArticle_OnClick;
  if Assigned(GetControl('MNLPDISPO')) then
    TMenuItem(GetControl('MNLPDISPO')).OnClick := MnLPDispo_OnClick;
  if Assigned(GetControl('MNLPREFORIGINE')) then
    TMenuItem(GetControl('MNLPREFORIGINE')).OnClick := MnLPRefOrigine_OnClick;
end ;

procedure TOM_STKMOUVEMENT.OnClose ;
begin
  Inherited ;

  PmFlux.Free;
end ;

procedure TOM_STKMOUVEMENT.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_STKMOUVEMENT.CalculField(FieldName: string);
begin
  inherited;
end;

procedure TOM_STKMOUVEMENT.ControlField(FieldName: string);
begin
  DisableControl;
  try
    if FieldName = '...' then
    begin
    end
  finally
    EnableControl;
  end;

  inherited;

  if LastError <> 0 then
    LastErrorMsg := TexteMessage[LastError];
end;

function TOM_STKMOUVEMENT.RecordIsValid: boolean;
begin
  Result := false;

  { Contrôle bas niveau }
  if not inherited RecordIsvalid then exit;

  { Contrôle écran }
  if Assigned(ecran) then
  begin
  end;

  { Contrôle écran + tob }
  if LastError = 0 then
  begin

  end;

  Result := LastError = 0;
  if LastError > 0 then
  begin
    LastErrorMsg := TexteMessage[LastError];
    if Assigned(ecran) then
    begin
      Case LastError of
        ErrQualifMvt : SetFocusControl('GSM_QUALIFMVT');
      end;
    end
    else
    begin
      fTob.AddChampSupValeur('Error', LastErrorMsg, false);
      if V_Pgi.Sav then PgiInfo(LastErrorMsg, 'wStkMouvement_Tom.RecordIsValid');
    end;
  end;
end;

procedure TOM_STKMOUVEMENT.SetStkTypeMvt;
begin
  SetString('GSM_STKTYPEMVT', GetFieldFromGSN('GSN_STKTYPEMVT', GetString('GSM_QUALIFMVT')));
end;

procedure TOM_STKMOUVEMENT.SetStkConvUnit; { GPAO STKMOUV }
var aUniteArticle : MyArrayValue;
begin
  if not assigned(ecran) then exit;

  if ecran.Name = 'GSMPHYSIQUE_FIC' then
  begin
    aUniteArticle := WGetSqlFieldsValues(['GA_QUALIFUNITESTO'], 'ARTICLE', 'GA_CODEARTICLE = "'+getString('GSM_CODEARTICLE')+'"');
    SetDouble('GSM_PHYSIQUE',ConvUnite(GetDouble('GSM_PHYSIQUE'),getControltext('UNITEFLUX'),aUniteArticle[0],'','',''));
    SetControltext('UNITEFLUX', aUniteArticle[0]);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Mise à jour en + des compteurs GQ
Mots clefs ... :
*****************************************************************}
procedure TOM_STKMOUVEMENT.AddToGQD;
begin
  if Pos(GetString('GSM_STKTYPEMVT'), 'PHY; CET') > 0 then
    DispoDetail.AddToGQD(GetCleGQD(false), GetPhysique(false))
end;


{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Mise à jour en + des compteurs GQ
Mots clefs ... :
*****************************************************************}
procedure TOM_STKMOUVEMENT.AddToGQ;
var
  iCSK  : integer;
  TobGSN: Tob;
begin
  { Tob système }
  GetTobGSN;
  GetTobCSK;

  TobGSN := VH_GP.TobGSN.FindFirst(['GSN_QUALIFMVT'], [GetString('GSM_QUALIFMVT')], false);
  if Assigned(TobGSN) then
  begin
    for iCSK := 0 to VH_Gp.TobCSK.Detail.Count-1 do
    begin
      if (VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE') <> '' ) and (Pos(VH_Gp.TobCSK.Detail[iCSK].G('CO_CODE'), TobGSN.G('GSN_QTEPLUS')) > 0) then
      begin
        if      GetString('GSM_STKTYPEMVT') = 'PHY' then Dispo.AddToGQ(GetCleGQ(false), VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE'), GetPhysique(false))
        else if GetString('GSM_STKTYPEMVT') = 'RES' then Dispo.AddToGQ(GetCleGQ(false), VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE'), GetStkPrevu(false))
        else if GetString('GSM_STKTYPEMVT') = 'ATT' then Dispo.AddToGQ(GetCleGQ(false), VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE'), GetStkPrevu(false))
        else if GetString('GSM_STKTYPEMVT') = 'CST' then { Rien }
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Mise à jour en - des compteurs GQ
Mots clefs ... :
*****************************************************************}
procedure TOM_STKMOUVEMENT.RemoveFromGQ(AvantModif: boolean);
var
  iCSK  : integer;
  TobGSN: Tob;

  function GetStkTypeMvt: String;
  begin
    if AvantModif then
      Result := GetStringAvantModif('GSM_STKTYPEMVT')
    else
      Result := GetString('GSM_STKTYPEMVT')
  end;
begin
  { Tob système }
  GetTobGSN;
  GetTobCSK;

  TobGSN := VH_GP.TobGSN.FindFirst(['GSN_QUALIFMVT'], [GetString('GSM_QUALIFMVT')], false);
  if Assigned(TobGSN) then
  begin
    for iCSK := 0 to VH_Gp.TobCSK.Detail.Count-1 do
    begin
      if (VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE') > '' ) and (Pos(VH_Gp.TobCSK.Detail[iCSK].G('CO_CODE'), TobGSN.G('GSN_QTEPLUS')) > 0) then
      begin
        if      GetString('GSM_STKTYPEMVT') = 'PHY' then Dispo.RemoveFromGQ(GetCleGQ(AvantModif), VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE'), GetPhysique(AvantModif))
        else if GetString('GSM_STKTYPEMVT') = 'RES' then Dispo.RemoveFromGQ(GetCleGQ(AvantModif), VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE'), GetStkPrevu(AvantModif))
        else if GetString('GSM_STKTYPEMVT') = 'ATT' then Dispo.RemoveFromGQ(GetCleGQ(AvantModif), VH_Gp.TobCSK.Detail[iCSK].G('CO_ABREGE'), GetStkPrevu(AvantModif))
        else if GetString('GSM_STKTYPEMVT') = 'CST' then { Rien }
      end;
    end;
  end;
end;

function TOM_STKMOUVEMENT.GetStkPrevu(AvantModif: boolean): Double;
begin
  if AvantModif then
    Result := iif(GetStringAvantModif('GSM_ETATMVT') = 'SOL', 0, Max(0, Abs(GetDoubleAvantModif('GSM_QPREVUE')) - Abs(GetDoubleAvantModif('GSM_PHYSIQUE'))))
  else
    Result := iif(GetString          ('GSM_ETATMVT') = 'SOL', 0, Max(0, Abs(GetDouble          ('GSM_QPREVUE')) - Abs(GetDouble          ('GSM_PHYSIQUE'))))
end;

function TOM_STKMOUVEMENT.Getphysique(AvantModif: Boolean): Double;
begin
  if AvantModif then
    Result := GetDoubleAvantModif('GSM_PHYSIQUE')
  else
    Result := GetDouble('GSM_PHYSIQUE')
end;

procedure TOM_STKMOUVEMENT.SetEtatMvt;
begin
  if Pos(GetString('GSM_STKTYPEMVT'), 'ATT;RES') > 0 then
  begin
    if Abs(GetDouble('GSM_PHYSIQUE')) >= Abs(GetDouble('GSM_QPREVUE')) then
      SetString('GSM_ETATMVT', 'SOL')
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 29/08/2003
Modifié le ... :   /  /
Description .. : Un mouvement ATT est toujours positif
Suite ........ : Un mouvement RES est toujours négatif
Mots clefs ... :
*****************************************************************}
procedure TOM_STKMOUVEMENT.SetSigne;
begin
  if      GetString('GSM_STKTYPEMVT') = 'ATT' then SetDouble('GSM_QPREVUE', +Abs(GetDouble('GSM_QPREVUE')))
  else if GetString('GSM_STKTYPEMVT') = 'RES' then SetDouble('GSM_QPREVUE', -Abs(GetDouble('GSM_QPREVUE')))
end;

function TOM_STKMOUVEMENT.GetCleGQD(AvantModif: boolean): tCleGQD;
begin
  if AvantModif then
  begin
    Result.Article        := GetStringAvantModif('GSM_ARTICLE');
    Result.Depot          := GetStringAvantModif('GSM_DEPOT');
    Result.StatutDispo    := GetStringAvantModif('GSM_STATUTDISPO');
    Result.StatutFlux     := GetStringAvantModif('GSM_STATUTFLUX');
    Result.Emplacement    := GetStringAvantModif('GSM_EMPLACEMENT');
    Result.LotInterne     := GetStringAvantModif('GSM_LOTINTERNE');
    Result.SerieInterne   := GetStringAvantModif('GSM_SERIEINTERNE');
    Result.RefProprio     := GetStringAvantModif('GSM_REFPROPRIO');
    Result.IndiceArticle  := GetStringAvantModif('GSM_INDICEARTICLE');
    Result.Marque         := GetStringAvantModif('GSM_MARQUE');
    Result.ChoixQualite   := GetStringAvantModif('GSM_CHOIXQUALITE');
    Result.RefAffectation := GetStringAvantModif('GSM_REFAFFECTATION');
  end
  else
  begin
    Result.Article        := GetString('GSM_ARTICLE');
    Result.Depot          := GetString('GSM_DEPOT');
    Result.StatutDispo    := GetString('GSM_STATUTDISPO');
    Result.StatutFlux     := GetString('GSM_STATUTFLUX');
    Result.Emplacement    := GetString('GSM_EMPLACEMENT');
    Result.LotInterne     := GetString('GSM_LOTINTERNE');
    Result.SerieInterne   := GetString('GSM_SERIEINTERNE');
    Result.RefProprio     := GetString('GSM_REFPROPRIO');
    Result.IndiceArticle  := GetString('GSM_INDICEARTICLE');
    Result.Marque         := GetString('GSM_MARQUE');
    Result.ChoixQualite   := GetString('GSM_CHOIXQUALITE');
    Result.RefAffectation := GetString('GSM_REFAFFECTATION');
  end;
end;

function TOM_STKMOUVEMENT.GetCleGQ(AvantModif: boolean): tCleGQ;
begin
  if AvantModif then
  begin
    Result.Article := GetStringAvantModif('GSM_ARTICLE');
    Result.Depot := GetStringAvantModif('GSM_DEPOT');
  end
  else
  begin
    Result.Article := GetString('GSM_ARTICLE');
    Result.Depot := GetString('GSM_DEPOT');
  end;
  Result.Cloture := false;
  Result.DateCloture := iDate1900;
end;

procedure TOM_STKMOUVEMENT.MnLpArticle_OnClick(Sender: TObject);
  function GetRange: string;
  begin
    Result := 'GA_ARTICLE=' + GetString('GSM_ARTICLE');
  end;
begin
	wCallGA(GetRange);
  RefreshDB;
end;

procedure TOM_STKMOUVEMENT.MnLpDispo_OnClick(Sender: TObject);
begin

//	CallMulGQ(GetCleGQ, '');
  RefreshDB;
end;

procedure TOM_STKMOUVEMENT.MnLpRefOrigine_OnClick(Sender: TObject);
begin

end;


procedure TOM_STKMOUVEMENT.ChangeFlux(PmFlux : TPopupMenuFlux);
var   sUniteFlux : String;
  aQualifMesure : MyArrayValue;
begin  // Vois si moyen d'optimiser ça : éviter de tester des valeurs en dur
  If PmFLux.Flux = 'STO' Then sUniteFlux := caract_art.QualifUniteSto
  Else If PmFLux.Flux = 'VTE' Then sUniteFlux := caract_art.UniteQteVte
  Else If PmFLux.Flux = 'ACH' Then sUniteFlux := caract_art.UniteQteAch
  Else If PmFLux.Flux = 'PRO' Then sUniteFlux := caract_art.UniteProd
  Else If PmFLux.Flux = 'CON' Then sUniteFlux := caract_art.UniteConso
  Else sUniteFlux := caract_art.QualifUniteSto;
  If sUniteFlux <> '' then // Si la fiche est appellée directement depuis le menu, on n'a pas encore de code article lors du premier passage ici => plantages dans les lignes suivantes si non protégées.
  Begin
    SetDouble('GSM_PHYSIQUE',ConvUnite(GetDouble('GSM_PHYSIQUE'),getControltext('UNITEFLUX'),sUniteFlux,'','',''));
    aQualifMesure := WGetSqlFieldsValues(['GME_QUALIFMESURE'], 'MEA', 'GME_MESURE = "'+sUniteFlux+'"');
    SetControlProperty('UNITEFLUX','PLUS','GME_QUALIFMESURE="'+aQualifMesure[0]+'"');
    SetControlText('UNITEFLUX', sUniteFlux);
  end;
end;

procedure TOM_STKMOUVEMENT.MNFlux_OnClick(Sender: TObject);
begin
  PmFLux.Flux := StringReplace(TPopupMenu(Sender).Name, 'MN', '', [rfIgnoreCase]);
  ChangeFlux(PmFlux);

  {for i := 0 to DS.FieldCount - 1 do
    CalculField(DS.Fields[i].FieldName);}

end;


function TOM_STKMOUVEMENT.GetUniteFromFlux: String;
begin
  Result := GetArgumentValue(stArgument, 'UNITE' + PmFLux.Flux);
end;

procedure TOM_STKMOUVEMENT.SetArticle;
begin
  if GetString('GSM_CODEARTICLE') <> '' then
  	SetString('GSM_ARTICLE', wGetArticleFromCodeArticle(GetString('GSM_CODEARTICLE')));
end;


procedure TOM_STKMOUVEMENT.SetCodeArticle;
begin
  if GetString('GSM_ARTICLE') <> '' then
    SetString('GSM_CODEARTICLE', wGetCodeArticleFromArticle(GetString('GSM_ARTICLE')));
end;

{$ENDIF}
Initialization
  registerclasses ( [ TOM_STKMOUVEMENT ] ) ;
end.

