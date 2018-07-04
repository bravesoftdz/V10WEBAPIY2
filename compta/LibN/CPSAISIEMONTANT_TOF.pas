{***********UNITE*************************************************
Auteur  ...... : YMO
Créé le ...... : 02/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSAISIEMONTANT ()
Mots clefs ... : TOF;CPSAISIEMONTANT
*****************************************************************}
Unit CPSAISIEMONTANT_TOF ;

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
     UTOF,
     Vierge ; 

Type
  TOF_CPSAISIEMONTANT = Class (TOF)
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

procedure TOF_CPSAISIEMONTANT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEMONTANT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEMONTANT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEMONTANT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEMONTANT.OnArgument (S : String ) ;
begin
  Inherited ;
  THNumEdit(GetControl('FMontant',True)).Value := strtoint(S) ;
end ;

procedure TOF_CPSAISIEMONTANT.OnClose ;
begin
  TFVierge (Ecran).retour := GetControlText('FMontant');
end ;

procedure TOF_CPSAISIEMONTANT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEMONTANT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPSAISIEMONTANT ] ) ;
end.
