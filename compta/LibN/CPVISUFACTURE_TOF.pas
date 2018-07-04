{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/11/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPVISUFACTURE ()
Mots clefs ... : TOF;CPVISUFACTURE
*****************************************************************}
Unit CPVISUFACTURE_TOF ;

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
  TOF_CPVISUFACTURE = Class (TOF)
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

procedure TOF_CPVISUFACTURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUFACTURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUFACTURE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUFACTURE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUFACTURE.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUFACTURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUFACTURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUFACTURE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPVISUFACTURE ] ) ; 
end.

