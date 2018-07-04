{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/03/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BALBUDECSEG ()
Mots clefs ... : TOF;BALBUDECSEG
*****************************************************************}
Unit BALBUDECSEG_TOF ;

Interface

Uses StdCtrls, Controls, Classes, 
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db, dbtables, mul, FE_Main, QRS1,
{$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, Ent1;

procedure BalBudecSecGen;

Type
  TOF_BALBUDECSEG = Class (TOF)
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

procedure BalBudecSecGen;
begin
  AGLLanceFiche('CP', 'CPBALBUDECSEG', '', '', '');
end;

procedure TOF_BALBUDECSEG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEG.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEG.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEG.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEG.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEG.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BALBUDECSEG.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BALBUDECSEG ] ) ; 
end.
