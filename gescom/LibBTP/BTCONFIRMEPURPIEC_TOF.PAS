{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/10/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONFIRMEPURPIEC ()
Mots clefs ... : TOF;BTCONFIRMEPURPIEC
*****************************************************************}
Unit BTCONFIRMEPURPIEC_TOF ;

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
  TOF_BTCONFIRMEPURPIEC = Class (TOF)
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

procedure TOF_BTCONFIRMEPURPIEC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURPIEC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURPIEC.OnUpdate ;
begin
  Inherited ;
  OneTOB.putValue('RESULTAT','OK');
end ;

procedure TOF_BTCONFIRMEPURPIEC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURPIEC.OnArgument (S : String ) ;
begin
  Inherited ;
  OneTOB := LaTOB;
end ;

procedure TOF_BTCONFIRMEPURPIEC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURPIEC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMEPURPIEC.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTCONFIRMEPURPIEC ] ) ; 
end.
