{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTCONCURRENTS_MUL ()
Mots clefs ... : TOF;RTCONCURRENTS_MUL
*****************************************************************}
Unit UTOFRTCONCURRENT_MUL;

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
  TOF_RTCONCURRENT_MUL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Function RTLanceFiche_RtConcurrent_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

Function RTLanceFiche_RtConcurrent_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_RTCONCURRENT_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
{$Ifdef GIGI}
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 (Ecran).HelpContext:=120000338 ;  //mcd 21/08/07 aide # GI
{$endif}
end ;


Initialization
  registerclasses ( [ TOF_RTCONCURRENT_MUL ] ) ;
end.
