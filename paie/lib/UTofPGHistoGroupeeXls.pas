{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTOXLS ()
Mots clefs ... : TOF;PGHISTOXLS
*****************************************************************}
Unit UTofPGHistoGroupeeXls ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     Vierge,
     UTOF ;

Type
  TOF_PGHISTOGROUPEEXLS = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PGHISTOGROUPEEXLS.OnUpdate ;
begin
  Inherited ;
  TFVierge(Ecran).Retour := GetControlText('LEFICHIER');
end ;

procedure TOF_PGHISTOGROUPEEXLS.OnArgument (S : String ) ;
begin
  Inherited ;
  TFVierge(Ecran).Retour := '';
end ;

Initialization
  registerclasses ( [ TOF_PGHISTOGROUPEEXLS ] ) ;
end.

