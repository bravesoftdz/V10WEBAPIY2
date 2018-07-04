unit UtofCubePerspectives;

interface

uses  Classes,forms,UtilGC,
      HEnt1,UTOF,
{$IF Defined(GIGI) or Defined(GRCLIGHT)}
      EntGC,
{$IFEND}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
      UtilSelection,UtilRT;

Type
{$ifdef AFFAIRE}
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
     TOF_CubePerspectives = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
    TOF_CubePerspectives = Class (TOF)
{$endif}
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad ; override ;
     END;

implementation
uses ParamSoc;
procedure TOF_CubePerspectives.OnArgument(stArgument : String );
var F : TForm;
begin
inherited ;
F := TForm (Ecran);
MulCreerPagesCL(F,'NOMFIC=GCTIERS');
if GetParamSocSecur('SO_RTGESTINFOS00V',False) then
  MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then
    begin
    SetControlVisible ('YTC_RESSOURCE1',false);
    SetControlVisible ('YTC_RESSOURCE2',false);
    SetControlVisible ('YTC_RESSOURCE3',false);
    SetControlVisible ('TYTC_RESSOURCE1',false);
    SetControlVisible ('TYTC_RESSOURCE2',false);
    SetControlVisible ('TYTC_RESSOURCE3',false);
    SetControlVisible ('T_MOISCLOTURE',false);
    SetControlVisible ('T_MOISCLOTURE_',false);
    SetControlVisible ('TT_MOISCLOTURE',false);
    SetControlVisible ('TT_MOISCLOTURE_',false);
    end
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 if (GetControl('RPE_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    end;
 if (GetControl('RPE_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
 if (GetControl('TRPE_REPRESENTANT') <> nil) then  SetControlVisible('TRPE_REPRESENTANT',false);
 if (GetControl('RPE_REPRESENTANT') <> nil) then  SetControlVisible('RPE_REPRESENTANT',false);
 if (GetControl('YTC_REPRESENTANT2') <> nil) then  SetControlVisible('YTC_REPRESENTANT2',false);
 if (GetControl('YTC_REPRESENTANT2_') <> nil) then  SetControlVisible('YTC_REPRESENTANT2_',false);
 if (GetControl('YTC_REPRESENTANT3') <> nil) then  SetControlVisible('YTC_REPRESENTANT3',false);
 if (GetControl('YTC_REPRESENTANT3_') <> nil) then  SetControlVisible('YTC_REPRESENTANT3_',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText ('TRPE_TYPETIERS', TraduireMemoire('Nature tiers'));
 SetControlProperty ('RPE_TYPETIERS', 'Complete', true);
 SetControlProperty ('RPE_TYPETIERS', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('RPE_TYPETIERS', 'Plus', VH_GC.AfNatTiersGRCGI);
 // STR 24/06/2008 FQ 10773
 SetControlText ('TT_NATUREAUXI', TraduireMemoire('Nature tiers'));
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 //
{$endif}
{$IFDEF GRCLIGHT}
  if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
    begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
{$ENDIF GRCLIGHT}

end;

procedure TOF_CubePerspectives.OnLoad;
begin
inherited;
SetControlText('XX_WHERE',RTXXWhereConfident('CON')) ;
end;

Initialization
registerclasses([TOF_CubePerspectives]);

end.
