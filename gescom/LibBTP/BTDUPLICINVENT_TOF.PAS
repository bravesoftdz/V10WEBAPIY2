{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/06/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDUPLICINVENT ()
Mots clefs ... : TOF;BTDUPLICINVENT
*****************************************************************}
Unit BTDUPLICINVENT_TOF ;

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
  TOF_BTDUPLICINVENT = Class (TOF)
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

procedure TOF_BTDUPLICINVENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDUPLICINVENT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDUPLICINVENT.OnUpdate ;
begin
  Inherited ;
  if ThEdit(getControl('INTITULE')).Text = '' then
  begin
    PGIInfo('Veuillez renseigner un intitulé pour cette liste');
    Exit;
  end;
  if ThEdit(getControl('RACINECODE')).Text = '' then
  begin
    PGIInfo('Veuillez renseigner une racine de code');
    Exit;
  end;
  LaTOB.SetString('OKOK','X');
end ;

procedure TOF_BTDUPLICINVENT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDUPLICINVENT.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTDUPLICINVENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDUPLICINVENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDUPLICINVENT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTDUPLICINVENT ] ) ; 
end.
