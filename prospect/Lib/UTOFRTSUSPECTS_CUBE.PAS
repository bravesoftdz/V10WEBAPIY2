{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTSUSPECTS_CUBE ()
Mots clefs ... : TOF;RTSUSPECTS_CUBE
*****************************************************************}
Unit UTOFRTSUSPECTS_CUBE ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     Fe_Main,
{$else}
     MainEAGL,
{$ENDIF}
     UTOF ;

Type
  TOF_RTSUSPECTS_CUBE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Function RTLanceFiche_RtSuspects_Cube(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

Function RTLanceFiche_RtSuspects_Cube(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_RTSUSPECTS_CUBE.OnArgument (S : String ) ;
begin
  Inherited ;
{$Ifdef GIGI}
 if (GetControl('RSU_REPRESENTANT') <> nil) then  SetControlVisible('RSU_REPRESENTANT',false);
 if (GetControl('TRSU_REPRESENTANT') <> nil) then  SetControlVisible('TRSU_REPRESENTANT',false);
 if (GetControl('RSU_ZONECOM') <> nil) then  SetControlVisible('RSU_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
{$endif}
end ;


Initialization
  registerclasses ( [ TOF_RTSUSPECTS_CUBE ] ) ; 
end.
