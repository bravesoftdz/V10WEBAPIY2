unit UTofSuspect_Mul;

interface

uses HEnt1, hmsgbox,UtilGc,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     FE_Main,
{$ENDIF}
{$ifdef AFFAIRE}
    UtofAftraducChampLibre,
{$ENDIF}
     Ent1, UTOF, Classes
     ,EntRT,ParamSoc ;

Function RTLanceFiche_Suspect_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type

{$ifdef AFFAIRE}
     TOF_Suspect_Mul = Class (TOF_AFTRADUCCHAMPLIBRE)  //mcd 19/06/06 pour affectation table libre ressource dans table libre suspect
{$else}
     TOF_Suspect_Mul = Class (TOF)
{$endif}
     public
        procedure OnArgument(Arguments : String ); override ;
    END ;

implementation

uses
  TiersUtil;

Function RTLanceFiche_Suspect_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Suspect_Mul.OnArgument(Arguments : String );
begin
	fMulDeTraitement := true;
  inherited ;
fTableName := 'SUSPECTS';
  if ( Not AutoriseCreationTiers ('PRO') ) or
     ( (VH_RT.RTExisteConfident=false) and (GetParamsocSecur('SO_RTCONFIDENTIALITE',False) = true) ) then
     SetControlVisible('BPROSPECT',False) ;
{$ifdef GIGI}
 SetControlVisible('RSU_ZONECOM',false);
 SetControlVisible('TRSU_ZONECOM',false);
 SetControlVisible('RSU_REPRESENTANT',false);
 SetControlVisible('TRSU_REPRESENTANT',false);
{$endif}
  //FQ10467 gestion des commerciaux
  if not GereCommercial then
  begin
    SetControlVisible ('RSU_REPRESENTANT', False);
    SetControlVisible ('TRSU_REPRESENTANT', False);
  end;
end;


Initialization
registerclasses([TOF_Suspect_Mul]);
end.
