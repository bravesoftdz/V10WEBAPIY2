{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/10/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONFIRMEPURCHA ()
Mots clefs ... : TOF;BTCONFIRMEPURCHA
*****************************************************************}
Unit BTCONFIRMEPURCHA_TOF ;

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
  TOF_BTCONFIRMEPURCHA = Class (TOF)
  private
  	OneTOB : TOB;
  public
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

procedure TOF_BTCONFIRMEPURCHA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURCHA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURCHA.OnUpdate ;
begin
  Inherited ;
  OneTOB.putValue('RESULTAT','OK');
end ;

procedure TOF_BTCONFIRMEPURCHA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURCHA.OnArgument (S : String ) ;
begin
  Inherited ;
  OneTOB := LaTOB;
end ;

procedure TOF_BTCONFIRMEPURCHA.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURCHA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURCHA.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTCONFIRMEPURCHA ] ) ; 
end.
