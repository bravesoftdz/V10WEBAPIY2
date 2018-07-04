{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CHOIXREGROUPEMENT ()
Mots clefs ... : TOF;CHOIXREGROUPEMENT
*****************************************************************}
Unit CHOIXREGROUPEMENT_TOF ;

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
  TOF_CHOIXREGROUPEMENT = Class (TOF)
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

procedure TOF_CHOIXREGROUPEMENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXREGROUPEMENT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXREGROUPEMENT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXREGROUPEMENT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXREGROUPEMENT.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXREGROUPEMENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXREGROUPEMENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXREGROUPEMENT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CHOIXREGROUPEMENT ] ) ; 
end.
