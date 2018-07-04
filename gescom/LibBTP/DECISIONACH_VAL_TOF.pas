{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DECISIONACH_VAL ()
Mots clefs ... : TOF;DECISIONACH_VAL
*****************************************************************}
Unit DECISIONACH_VAL_TOF ;

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
     HMsgBox,
     UTOB,
     UTOF ;

Type
  TOF_DECISIONACH_VAL = Class (TOF)
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

procedure TOF_DECISIONACH_VAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_VAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_VAL.OnUpdate ;
begin
  Inherited ;
  if TCheckBox(getControl('CBECLATCDEFOU')).Checked then LaTOB.PutValue('ECLAT','X')
  																									else LaTOB.PutValue('ECLAT','-');
  laTOB.putValue('DELAISECURITE',THNumEdit(GetCOntrol('DELAISECURITE')).Value);
  laTOB.putValue('OK','X');
end ;

procedure TOF_DECISIONACH_VAL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_VAL.OnArgument (S : String ) ;
begin
  Inherited ;
  TCheckBox(getControl('CBECLATCDEFOU')).Checked := (laTOB.GetValue('ECLAT')='X');
end ;

procedure TOF_DECISIONACH_VAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_VAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACH_VAL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_DECISIONACH_VAL ] ) ; 
end.

