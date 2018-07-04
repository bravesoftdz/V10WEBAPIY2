{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/02/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMODIFLISTE ()
Mots clefs ... : TOF;BTMODIFLISTE
*****************************************************************}
Unit BTMODIFLISTE_TOF ;

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
     HMsgBox, 
     UTOF ; 

Type
  TOF_BTMODIFLISTE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_BTMODIFLISTE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFLISTE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFLISTE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFLISTE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFLISTE.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFLISTE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFLISTE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFLISTE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTMODIFLISTE ] ) ; 
end.
