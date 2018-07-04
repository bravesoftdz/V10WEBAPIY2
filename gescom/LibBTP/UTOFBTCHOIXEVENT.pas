{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/06/2006
Modifié le ... :   /  / 
Description .. : Source TOF de la FICHE : BTSUPPEVENT ()
Mots clefs ... : TOF;BTCHOIXEVENT
*****************************************************************}
Unit UTOFBTCHOIXEVENT ;

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
     AglInit,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_BTCHOIXEVENT = Class (TOF)
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

procedure TOF_BTCHOIXEVENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXEVENT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXEVENT.OnUpdate ;
begin
  Inherited ;

	If TRadioButton(GetControl('APPEL')).Checked then
     LaTob.PutValue('RETOUR', '1')
  Else If TRadioButton(GetControl('INTERVENTION')).Checked then
     LaTob.PutValue('RETOUR', '2')
  Else If TRadioButton(GetControl('ACTION_GRC')).Checked then
     LaTob.PutValue('RETOUR', '3')
  Else If TRadioButton(GetControl('PARCMAT')).Checked then
     LaTob.PutValue('RETOUR', '4')
  Else If TRadioButton(GetControl('CHANTIER')).Checked then
     LaTob.PutValue('RETOUR', '5');

	TheTob := LaTob;

end ;

procedure TOF_BTCHOIXEVENT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXEVENT.OnArgument (S : String ) ;
begin
  Inherited ;

  if  Latob.GetValue('VAPPEL')        = '1' then
    TRadioButton(GetControl('APPEL')).Visible        := True
  else
    TRadioButton(GetControl('APPEL')).Visible        := False;
  //
  if  Latob.GetValue('VINTERVENTION') = '1' then
    TRadioButton(GetControl('INTERVENTION')).Visible := True
  else
    TRadioButton(GetControl('INTERVENTION')).Visible := False;

  if  Latob.GetValue('VACTION_GRC')   = '1' then
    TRadioButton(GetControl('ACTION_GRC')).Visible   := True
  else
    TRadioButton(GetControl('ACTION_GRC')).Visible   := False;

  if  Latob.GetValue('VPARCMAT')      = '1' then
    TRadioButton(GetControl('PARCMAT')).Visible      := True
  else
    TRadioButton(GetControl('PARCMAT')).Visible      := False;

  if  Latob.GetValue('VCHANTIER')     = '1' then
    TRadioButton(GetControl('CHANTIER')).Visible     := True
  else
    TRadioButton(GetControl('CHANTIER')).Visible     := false;

  TRadioButton(GetControl('APPEL')).Checked := True;


end ;

procedure TOF_BTCHOIXEVENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXEVENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCHOIXEVENT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTCHOIXEVENT ] ) ;
end.


