{***********UNITE*************************************************
Auteur  ...... : GHA
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULHABILITATIONS ()
Mots clefs ... : TOF;MULHABILITATIONS
*****************************************************************}
Unit MULHABILITATIONS_TOF ;

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
  TOF_MULHABILITATIONS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    GrpUtiCmbbx : THValComboBox;
    procedure DoChanged(sender : TObject);
  end;

Implementation

procedure TOF_MULHABILITATIONS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULHABILITATIONS.OnArgument (S : String ) ;
begin
  Inherited ;
  // Init des ctrl
  SetControlEnabled('TPHL_USERGRP',FALSE);
  SetControlEnabled('PHL_USERGRP',FALSE);
  SetControlEnabled('TPHL_USER',FALSE);
  SetControlEnabled('PHL_USER',FALSE);

  GrpUtiCmbbx := THValComboBox(GetControl('PHL_GRPUTI',TRUE));
  if GrpUtiCmbbx <> nil then
    GrpUtiCmbbx.OnChange := DoChanged;
end ;

procedure TOF_MULHABILITATIONS.DoChanged(sender: TObject);
begin
  if (sender as THValComboBox).name = 'PHL_GRPUTI' then
  begin
    if GetControlText('PHL_GRPUTI') = 'GRP' then
    begin
      SetControlEnabled('TPHL_USER',FALSE);
      SetControlEnabled('PHL_USER',FALSE);
      if GetControlText('PHL_USER') <> '' then
        SetControlText('PHL_USER','');

      SetControlEnabled('TPHL_USERGRP',TRUE);
      SetControlEnabled('PHL_USERGRP',TRUE);
    end
    else if GetControlText('PHL_GRPUTI') = 'UTI' then
    begin
      SetControlEnabled('TPHL_USER',TRUE);
      SetControlEnabled('PHL_USER',TRUE);

      SetControlEnabled('TPHL_USERGRP',FALSE);
      SetControlEnabled('PHL_USERGRP',FALSE);
      if GetControlText('PHL_USERGRP') <> '' then
        SetControlText('PHL_USERGRP','');
    end
    else
    begin
      SetControlEnabled('TPHL_USERGRP',FALSE);
      SetControlEnabled('PHL_USERGRP',FALSE);
      SetControlEnabled('TPHL_USER',FALSE);
      SetControlEnabled('PHL_USER',FALSE);
      if GetControlText('PHL_USERGRP') <> '' then
        SetControlText('PHL_USERGRP','');
      if GetControlText('PHL_USER') <> '' then
        SetControlText('PHL_USER','');
    end;
  end;
end;

procedure TOF_MULHABILITATIONS.OnClose ;
begin
  Inherited ;
  GrpUtiCmbbx.Free;
end ;

procedure TOF_MULHABILITATIONS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULHABILITATIONS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULHABILITATIONS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULHABILITATIONS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULHABILITATIONS.OnUpdate ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MULHABILITATIONS ] ) ; 
end.
