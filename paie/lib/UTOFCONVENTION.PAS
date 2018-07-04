{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFCONVENTION ()
Mots clefs ... : TOF;UTOFCONVENTION
*****************************************************************}
Unit UTOFCONVENTION;

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
  TOF_UTOFCONVENTION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      CEG, STD, DOS: Boolean;
  end ;

Implementation
uses
  PgOutils;

procedure TOF_UTOFCONVENTION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCONVENTION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCONVENTION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCONVENTION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCONVENTION.OnArgument (S : String ) ;
begin
  Inherited ;
  AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1}
   STD := TRUE;
{$ENDIF}
  
  if (CEG = FALSE) and (STD = FALSE) then
    SetControlEnabled('BInsert', False);
end ;

procedure TOF_UTOFCONVENTION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCONVENTION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_UTOFCONVENTION.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_UTOFCONVENTION ] ) ; 
end.
