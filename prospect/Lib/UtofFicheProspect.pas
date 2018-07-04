unit UtofFicheProspect;

interface

uses  Classes,forms,
      HEnt1,UTOF,Paramsoc,UtilGC,
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     Fe_Main,
{$ENDIF}
{$ifdef GIGI}
      EntGC,
{$ENDIF GIGI}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
     UtilRT,UtilSelection;

Type
{$ifdef AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
    TOF_FicheProspect = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
    TOF_FicheProspect = Class (TOF)
{$endif}
     Private
        stProduitpgi : string;
     public
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad ; override ;
     END;

Function RTLanceFiche_FicheProspect(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

Function RTLanceFiche_FicheProspect(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_FicheProspect.OnLoad;
var Confid : string;
begin
inherited;
  if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';
  SetControlText('XX_WHERE',RTXXWhereConfident(Confid)) ;
end;

procedure TOF_FicheProspect.OnArgument(stArgument : String );
var    F : TForm;
begin
	fMulDeTraitement := true;
inherited ;
fTableName := 'TIERS';

F := TForm (Ecran);
if stArgument = 'GRF' then
  begin
  if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
      MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
  end
else MulCreerPagesCL(TForm (Ecran),'NOMFIC=GCTIERS');
stProduitpgi := stArgument;
if stProduitpgi = '' then stProduitpgi := 'GRC';
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
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
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
{$endif}
end;

Initialization
registerclasses([TOF_FicheProspect]);
end.
