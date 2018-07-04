{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/04/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMULAFFAIREINTER ()
Mots clefs ... : TOF;BTMULAFFAIREINTER
*****************************************************************}
Unit BTMULAFFAIREINTER_TOF ;

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
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_BTMULAFFAIREINTER = Class (TOF)
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

procedure TOF_BTMULAFFAIREINTER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFAIREINTER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFAIREINTER.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFAIREINTER.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFAIREINTER.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFAIREINTER.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFAIREINTER.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMULAFFAIREINTER.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTMULAFFAIREINTER ] ) ;
end.
