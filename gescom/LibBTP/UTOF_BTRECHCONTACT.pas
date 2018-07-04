{***********UNITE*************************************************
Auteur  ...... : F.V 
Créé le ...... : 18/10/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECHCONTACT ()
Mots clefs ... : TOF;BTRECHCONTACT
*****************************************************************}
Unit UTOF_BTRECHCONTACT ;

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
     UTOF ; 

Type
  TOF_BTRECHCONTACT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function OkRechercheContact (ZoneSTock : TControl; ZoneParam : array of string ; Requete : string) : boolean;

Implementation

function OkRechercheContact (ZoneSTock : TControl; ZoneParam : array of string ; Requete : string) : boolean;
begin

	result := false;

end;

procedure TOF_BTRECHCONTACT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHCONTACT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHCONTACT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHCONTACT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHCONTACT.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHCONTACT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHCONTACT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHCONTACT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTRECHCONTACT ] ) ; 
end.

