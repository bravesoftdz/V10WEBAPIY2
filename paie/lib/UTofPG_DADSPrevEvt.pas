{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/12/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DADS_PREVEVT ()
Mots clefs ... : TOF;PAIE;PGDADSU
*****************************************************************}
Unit UTofPG_DADSPrevEvt;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
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
     Vierge,
     HMsgBox, 
     UTOF ; 

Type
  TOF_PGDADSPREVEVT = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/12/2007
Modifié le ... :   /  /    
Description .. : OnUpdate
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGDADSPREVEVT.OnUpdate;
begin
Inherited ;
TFVierge(Ecran).Retour:= GetControltext ('CODE');
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/12/2007
Modifié le ... :   /  /    
Description .. : OnArgument
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGDADSPREVEVT.OnArgument (S : String );
var
CodeEvt : String;
begin
Inherited ;
CodeEvt:= ReadTokenPipe (S, ';');
TFVierge(Ecran).Retour:= CodeEvt;
SetControlText ('CODE', CodeEvt);
end;

Initialization
registerclasses ([TOF_PGDADSPREVEVT]);
end.

