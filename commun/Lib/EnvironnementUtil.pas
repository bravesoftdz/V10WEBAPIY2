{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/04/2006
Modifié le ... :   /  /
Description .. : Chargement de données au fur et à mesure du besoin
Suite ........ : pour éviter de tout charger à la connexion
Mots clefs ... :
*****************************************************************}
unit EnvironnementUtil;

interface

uses EntGC
     , Utob
     , HCtrls
{$IFDEF EAGLCLIENT}
     , Utileagl
     , MaineAgl
{$ELSE EAGLCLIENT}
 {$IFNDEF DBXPRESS}
     , dbtables
 {$ELSE DBXPRESS}
     , uDbxDataSet
 {$ENDIF DBXPRESS}
{$ENDIF EAGLCLIENT}
     , HEnt1
     , SysUtils
     ;

function ImpactPiece_ChargeElement : TOB;
function ImpactPiece_ChargeTiers(LeFlux, LaDate : string) : TOB;

implementation

function ImpactPiece_ChargeElement : TOB;
var Qry : TQuery;
begin
  if VH_GC.TOBElementImpactPce.detail.count = 0 then
  begin
    Qry := OpenSQL('SELECT * FROM COMMUN WHERE CO_TYPE="GEF"', true);
    try
      VH_GC.TOBElementImpactPce.LoadDetailDB('COMMUN', '', 'CO_CODE', Qry, false);
    finally
      Ferme(Qry);
    end;
  end;
  Result := VH_GC.TOBElementImpactPce;
end;

function ImpactPiece_ChargeTiers(LeFlux, LaDate : string) : TOB;
var TobFlux : TOB;
    Sql, SigneDte : string;
    Qry : TQuery;
begin
  if LaDate = '' then
    LaDate := UsDateTime(iDate2099)
    else
    LaDate := UsDateTime(StrToDate(LaDate));
  if LaDate = UsDateTime(iDate2099) then
    SigneDte := '='
    else
    SigneDte := '>';
  if LeFlux = 'VEN' then
    TobFlux := VH_GC.TOBTiersImpactPce.detail[0]
  else if leFlux = 'ACH' then
    TobFlux := VH_GC.TOBTiersImpactPce.detail[1]
  else
    TobFlux := nil;
  if assigned(TobFLux) then
  begin
    if TobFlux.detail.count = 0 then
    begin
      Sql := 'SELECT * FROM TIERSIMPACTPIECE WHERE GTI_FLUX = "' + LeFlux + '"'
           + ' AND GTI_DATEPARAM = (SELECT DISTINCT(MIN(DT.GTI_DATEPARAM)) FROM TIERSIMPACTPIECE DT'
           + '                      WHERE DT.GTI_FLUX = "' + LeFlux + '"'
           + '                      AND DT.GTI_DATEPARAM' + SigneDte + '"' + LaDate + '")';
      Qry := OpenSQL(Sql, true);
      try
        TobFlux.LoadDetailDB('TIERSIMPACTPIECE', '', '', Qry, false);
      finally
        Ferme(Qry);
      end;
    end;
  end;
  Result := TobFlux;
end;

end.
