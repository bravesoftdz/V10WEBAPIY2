unit factdomaines;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB, EdtEtat, EdtDoc,
  EdtREtat, EdtRDoc,
  {$ENDIF}
  SysUtils, Dialogs, UtilPGI, AGLInit,
  Math, EntGC, Classes, HMsgBox,
  {$IFDEF CHR}HRReglement, {$ENDIF}
  {$IFDEF GRC}UtilRT,{$ENDIF}
  UtilGC, ParamSoc;

procedure InitDomaineAct;
procedure AppliqueCoefDomaineActOuv (TOBOUV : TOB; TypeAffaire : String='A'; RecupPrix : string = 'PUH');
procedure AppliqueCoefDomaineLig (TOBL : TOB; TypeAffaire : String='A');
procedure GetCoefDomaine (Domaine : string; var Coeffg,Coefmarg : double; TypeAffaire : String='A');

implementation
var TOBDOMAINES : TOB;

procedure InitDomaineAct;
begin
	TOBDomaines.clearDetail;
end;

procedure GetCoefDomaine (Domaine : string; var Coeffg,Coefmarg : double; TypeAffaire : String='A');
var QQ : TQuery;
    TOBDOMACT : TOB;
begin
  if Domaine = '' then exit;

  Coeffg := 0;
  Coefmarg := 0;

  TOBDOMACT := TOBDOMAINES.findFirst(['BTD_CODE'],[Domaine],true);

  if TOBDOMACT = nil then
  begin
    QQ := OpenSql('SELECT * FROM BTDOMAINEACT WHERE BTD_CODE="'+Domaine+'"',true,-1, '', True);
    if not QQ.eof then
    begin
      TOBDOMACT := TOB.create ('BTDOMAINEACT',TOBDOMAINES,-1);
      TOBDOMACT.SelectDB ('',QQ);
    end;
  end;
  if TOBDOMACT <> NIL then
  begin
    if TypeAffaire = 'W' then
    begin
      if TOBDOMACT.GetValue('BTD_COEFFG_APPEL') <> 0 then
      begin
        Coeffg := TOBDOMACT.GetDouble('BTD_COEFFG_APPEL');
      end;
      if TOBDOMACT.GetValue('BTD_COEFMARG_APPEL') <> 0 then
      begin
        CoefMarg := TOBDOMACT.GetDouble('BTD_COEFMARG_APPEL');
      end;
    end
    Else
    begin
      if TOBDOMACT.GetValue('BTD_COEFFG') <> 0 then
      begin
        Coeffg := TOBDOMACT.GetDouble('BTD_COEFFG');
      end;
      if TOBDOMACT.GetValue('BTD_COEFMARG') <> 0 then
      begin
        CoefMarg := TOBDOMACT.GetDouble('BTD_COEFMARG');
      end;
    end;
  end;
  ferme(QQ);
end;

procedure AppliqueCoefDomaineLig (TOBL : TOB; TypeAffaire : String='A');
var QQ : TQuery;
    TOBDOMACT : TOB;
begin

  if TOBL.getvalue('GL_DOMAINE')='' then exit;
  if GetInfoParPiece(TOBL.parent.getvalue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT') = 'ACH' then exit;

  TOBDOMACT := TOBDOMAINES.findFirst(['BTD_CODE'],[TOBL.getvalue('GL_DOMAINE')],true);

  if TOBDOMACT = nil then
  begin
    QQ := OpenSql('SELECT * FROM BTDOMAINEACT WHERE BTD_CODE="'+TOBL.getvalue('GL_DOMAINE')+'"',true,-1, '', True);
    if not QQ.eof then
    begin
      TOBDOMACT := TOB.create ('BTDOMAINEACT',TOBDOMAINES,-1);
      TOBDOMACT.SelectDB ('',QQ);
    end;
  end;

  //FV1 : 12/09/2013 - FS#33 - POUCHAIN : En facturation de consos, les coefficients de l'UO ne sont pas repris
  if TOBDOMACT <> NIL then
  begin
    if TypeAffaire ='W' then
    begin
      if TOBDOMACT.GetValue('BTD_COEFFG_APPEL') <> 0 then
      begin
        TOBL.putValue('GL_COEFFG',TOBDOMACT.GetValue('BTD_COEFFG_APPEL')-1);
      end;
      if TOBDOMACT.GetValue('BTD_COEFMARG_APPEL') <> 0 then
      begin
        TOBL.putValue('GL_COEFMARG',TOBDOMACT.GetValue('BTD_COEFMARG_APPEL'));
        TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
      end;
    end
    else
    begin
      if TOBDOMACT.GetValue('BTD_COEFFG') <> 0 then
      begin
        TOBL.putValue('GL_COEFFG',TOBDOMACT.GetValue('BTD_COEFFG')-1);
      end;
      if TOBDOMACT.GetValue('BTD_COEFMARG') <> 0 then
      begin
        TOBL.putValue('GL_COEFMARG',TOBDOMACT.GetValue('BTD_COEFMARG'));
        TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
      end;
    end;
  end;

  ferme(QQ);

end;

procedure AppliqueCoefDomaineActOuv (TOBOUV : TOB; TypeAffaire : String='A'; RecupPrix : string = 'PUH');
var QQ : TQuery;
    TOBDOMACT : TOB;
begin

  if TOBOUV.getvalue('BLO_DOMAINE')='' then exit;

  TOBDOMACT := TOBDOMAINES.findFirst(['BTD_CODE'],[TOBOUV.getvalue('BLO_DOMAINE')],true);

  if TOBDOMACT = nil then
  begin
    QQ := OpenSql('SELECT * FROM BTDOMAINEACT WHERE BTD_CODE="'+TOBOUV.getvalue('BLO_DOMAINE')+'"',true,-1, '', True);
    if not QQ.eof then
    begin
      TOBDOMACT := TOB.create ('BTDOMAINEACT',TOBDOMAINES,-1);
      TOBDOMACT.SelectDB ('',QQ);
    end;
  end;

  //FV1 : 12/09/2013 - FS#33 - POUCHAIN : En facturation de consos, les coefficients de l'UO ne sont pas repris
  if TOBDOMACT <> NIL then
  begin
    if TypeAffaire ='W' then
    begin
      if TOBDOMACT.GetValue('BTD_COEFFG_APPEL') <> 0 then
      begin
        if (RecupPrix = 'PUH') or (RecupPrix = 'DPR') then
        begin
        	TOBOUV.putValue('BLO_COEFFG',TOBDOMACT.GetValue('BTD_COEFFG_APPEL')-1);
        end;
      end;
      if TOBDOMACT.GetValue('BTD_COEFMARG_APPEL') <> 0 then
      begin
        if (RecupPrix = 'PUH') then
        begin
        	TOBOUV.putValue('BLO_COEFMARG',TOBDOMACT.GetValue('BTD_COEFMARG_APPEL'));
          TOBOuv.PutValue('POURCENTMARG',Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
      end;
    end
    Else
    Begin
      if TOBDOMACT.GetValue('BTD_COEFFG') <> 0 then
      begin
        if (RecupPrix = 'PUH') or (RecupPrix = 'DPR') then
        begin
        	TOBOUV.putValue('BLO_COEFFG',TOBDOMACT.GetValue('BTD_COEFFG')-1);
        end;
      end;
      if TOBDOMACT.GetValue('BTD_COEFMARG') <> 0 then
      begin
        if (RecupPrix = 'PUH') then
        begin
        	TOBOUV.putValue('BLO_COEFMARG',TOBDOMACT.GetValue('BTD_COEFMARG'));
          TOBOuv.PutValue('POURCENTMARG',Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
      end;
    end;
  end;

  ferme(QQ);

end;

INITIALIZATION
  TOBDOMAINES := TOB.Create ('LES DOMAINES',nil,-1);
FINALIZATION
  TOBDOMAINES.free;

end.

