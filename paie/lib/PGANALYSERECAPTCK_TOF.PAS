{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGANALYSERECAPTCK ()
Mots clefs ... : TOF;PGANALYSERECAPTCK
*****************************************************************}
Unit PGANALYSERECAPTCK_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_PGANALYSERECAPTCK = Class (TOF)
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

procedure TOF_PGANALYSERECAPTCK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSERECAPTCK.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSERECAPTCK.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSERECAPTCK.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSERECAPTCK.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSERECAPTCK.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSERECAPTCK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGANALYSERECAPTCK.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_PGANALYSERECAPTCK ] ) ; 
end.
