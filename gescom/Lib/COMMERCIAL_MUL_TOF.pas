{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : COMMERCIAL_MUL ()
Mots clefs ... : TOF;COMMERCIAL_MUL
*****************************************************************}
Unit COMMERCIAL_MUL_TOF ;

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
     Ent1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_COMMERCIAL_MUL = Class (TOF)
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

procedure TOF_COMMERCIAL_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_COMMERCIAL_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_COMMERCIAL_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_COMMERCIAL_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_COMMERCIAL_MUL.OnArgument (S : String ) ;
var CC : THValComboBox;
begin
  Inherited ;
   CC:=THValComboBox(GetControl('GCL_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
end ;

procedure TOF_COMMERCIAL_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_COMMERCIAL_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_COMMERCIAL_MUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_COMMERCIAL_MUL ] ) ; 
end.

