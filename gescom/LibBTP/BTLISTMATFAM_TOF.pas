{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/10/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTLISTMATFAM ()
Mots clefs ... : TOF;BTLISTMATFAM
*****************************************************************}
Unit BTLISTMATFAM_TOF ;

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
     HTB97,
     UTOF ; 

Type
  TOF_BTLISTMATFAM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private

    BTErase : TToolbarButton97;

    procedure OnErase(Sender : TObject);

  end ;

Implementation

procedure TOF_BTLISTMATFAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTLISTMATFAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTLISTMATFAM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTLISTMATFAM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTLISTMATFAM.OnArgument (S : String ) ;
begin
  Inherited ;

  BTErase         := TToolbarButton97(GetControl('BTERASE'));
  BTErase.OnClick := OnErase;
  
end ;

procedure TOF_BTLISTMATFAM.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTLISTMATFAM.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTLISTMATFAM.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BTLISTMATFAM.OnErase(sender : TObject);
begin
  if Assigned(GetControl('BMA_CODEFAMILLE'))    then SetControlText('BMA_CODEFAMILLE',  '');
  if Assigned(GetControl('BMA_CODEFAMILLE_'))   then SetControlText('BMA_CODEFAMILLE_', '');
  if Assigned(GetControl('BMA_CODEARTICLE'))    then SetControlText('BMA_CODEARTICLE',  '');
  if Assigned(GetControl('BMA_RESSOURCE'))      then SetControlText('BMA_RESSOURCE',    '');
end;

Initialization
  registerclasses ( [ TOF_BTLISTMATFAM ] ) ; 
end.

 