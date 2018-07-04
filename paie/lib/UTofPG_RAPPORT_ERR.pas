{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/02/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EDIT_RAPPORT_ERR ()
Mots clefs ... : TOF;EDIT_RAPPORT_ERR
*****************************************************************}
Unit UTofPG_RAPPORT_ERR;

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
  TOF_EDIT_RAPPORT_ERR = Class (TOF)
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

procedure TOF_EDIT_RAPPORT_ERR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_RAPPORT_ERR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_RAPPORT_ERR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_RAPPORT_ERR.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_RAPPORT_ERR.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_RAPPORT_ERR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_RAPPORT_ERR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_RAPPORT_ERR.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_EDIT_RAPPORT_ERR ] ) ; 
end.
