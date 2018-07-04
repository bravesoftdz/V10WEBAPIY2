{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/02/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EDIT_GARDE_DOS ()
Mots clefs ... : TOF;EDIT_GARDE_DOS
*****************************************************************}
Unit UTofPG_EDIT_GARDE_DOS;

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
  TOF_EDIT_GARDE_DOS = Class (TOF)
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

procedure TOF_EDIT_GARDE_DOS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_GARDE_DOS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_GARDE_DOS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_GARDE_DOS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_GARDE_DOS.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_GARDE_DOS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_GARDE_DOS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_GARDE_DOS.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ UTofPG_EDIT_GARDE_DOS ] ) ; 
end.
