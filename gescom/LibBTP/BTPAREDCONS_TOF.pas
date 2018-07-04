{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 29/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPAREDCONS ()
Mots clefs ... : TOF;BTPAREDCONS
*****************************************************************}
Unit BTPAREDCONS_TOF ;

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
     AglInit, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_BTPAREDCONS = Class (TOF)
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

procedure TOF_BTPAREDCONS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPAREDCONS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPAREDCONS.OnUpdate ;
begin
  Inherited ;
  if TCheckBox(GetCOntrol('CBDEMANDE')).State = CbChecked then LaTOB.PutValue('ALLLEVEL','X')
                                                          else LaTOB.PutValue('ALLLEVEL','-');
  if TCheckBox(GetCOntrol('CBSAUTPAGEDEST')).State = CbChecked then LaTOB.PutValue('SAUTPAGEDEST','X')
                                                               else LaTOB.PutValue('SAUTPAGEDEST','-');
  if TCheckBox(GetCOntrol('CBMODEEDIT')).State = CbChecked then LaTOB.PutValue('IMPGLOBALE','X')
                                                           else LaTOB.PutValue('IMPGLOBALE','-');
  TheTOB := LaTOB;
end ;

procedure TOF_BTPAREDCONS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPAREDCONS.OnArgument (S : String ) ;
begin
  Inherited ;
  if LaTOB <> nil then TCheckBox(GetCOntrol('CBDEMANDE')).checked := (LaTOB.getValue('ALLLEVEL')='X');
  if LaTOB <> nil then TCheckBox(GetCOntrol('CBSAUTPAGEDEST')).checked := (LaTOB.getValue('SAUTPAGEDEST')='X');
  if LaTOB <> nil then TCheckBox(GetCOntrol('CBMODEEDIT')).checked := (LaTOB.getValue('IMPGLOBALE')='X');
end ;

procedure TOF_BTPAREDCONS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPAREDCONS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPAREDCONS.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTPAREDCONS ] ) ; 
end.

