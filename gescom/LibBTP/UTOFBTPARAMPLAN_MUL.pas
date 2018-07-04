{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMPLAN_MUL ()
Mots clefs ... : TOF;BTPARAMPLAN_MUL
*****************************************************************}
Unit UTOFBTPARAMPLAN_MUL ;

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
  TOF_BTPARAMPLAN_MUL = Class (TOF)
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

procedure TOF_BTPARAMPLAN_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMPLAN_MUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTPARAMPLAN_MUL ] ) ; 
end.

