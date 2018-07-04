{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/06/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCHOIXFINSAISIE ()
Mots clefs ... : TOF;BTCHOIXFINSAISIE
*****************************************************************}
Unit BTCHOIXFINSAISIE_TOF ;

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
     UTOF
      ;

Type
  TOF_BTCHOIXFINSAISIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	OneTOb : TOB;
  end ;

Implementation

procedure TOF_BTCHOIXFINSAISIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXFINSAISIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXFINSAISIE.OnUpdate ;
begin
  Inherited ;
  if TRadioButton (GetCOntrol('RBSAUVESAISIE')).Checked then OneTOB.PutValue('RETOUR',0)
  																											else OneTOB.PutValue('RETOUR',1);
end ;

procedure TOF_BTCHOIXFINSAISIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXFINSAISIE.OnArgument (S : String ) ;
begin
  Inherited ;
  OneTOB := LaTOB;
end ;

procedure TOF_BTCHOIXFINSAISIE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXFINSAISIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXFINSAISIE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTCHOIXFINSAISIE ] ) ; 
end.
