{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTFICHESUSPECT ()
Mots clefs ... : TOF;RTFICHESUSPECT
*****************************************************************}
Unit UTOFRTFICHESUSPECT;

Interface

Uses  
     Classes,
{$IFNDEF EAGLCLIENT}
     Fe_Main,
{$else}
     MainEAGL,
{$ENDIF}
     UTOF ;

Type
  TOF_RTFICHESUSPECT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Function RTLanceFiche_RtFicheSuspect(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

Function RTLanceFiche_RtFicheSuspect(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_RTFICHESUSPECT.OnArgument (S : String ) ;
begin
  Inherited ;
{$Ifdef GIGI}
 if (GetControl('RSU_REPRESENTANT') <> nil) then  SetControlVisible('RSU_REPRESENTANT',false);
 if (GetControl('TRSU_REPRESENTANT') <> nil) then  SetControlVisible('TRSU_REPRESENTANT',false);
 if (GetControl('RSU_ZONECOM') <> nil) then  SetControlVisible('RSU_ZONECOM',false);
 if (GetControl('TRSU_ZONECOM') <> nil) then  SetControlVisible('TRSU_ZONECOM',false);
 if (GetControl('ISGI') <> nil) then  SetControlChecked('ISGI',true);
{$endif}
end ;


Initialization
  registerclasses ( [ TOF_RTFICHESUSPECT ] ) ; 
end.
