{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/10/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFICHEMATCONFIE ()
Mots clefs ... : TOF;BTFICHEMATCONFIE
*****************************************************************}
Unit BTFICHEMATCONFIE_TOF ;

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
  TOF_BTFICHEMATCONFIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private

    DateRestitution : ThEdit;
    DatedeRemise    : ThEdit;
    XX_WHERE        : THEdit;
    //
    BTErase     : TToolbarButton97;
    //
    RBRestitue    : TRadioButton;
    RBARestituer  : TRadioButton;
    RBTous        : TRadioButton;

    procedure OnErase(Sender : TObject);
    procedure OnRestitue(sender : TObject);
    procedure OnARecevoir(sender: TObject);
    procedure OnTous(sender: TObject);

  end ;

Implementation

procedure TOF_BTFICHEMATCONFIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFICHEMATCONFIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFICHEMATCONFIE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTFICHEMATCONFIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFICHEMATCONFIE.OnArgument (S : String ) ;
begin
  Inherited ;

  if assigned(GetControl('XX_WHERE')) then
  begin
    XX_Where            := THEdit(GetControl('XX_WHERE'));
  end;

  if assigned(GetControl('BTERASE')) then
  begin
    BTErase             := TToolbarButton97(GetControl('BTERASE'));
    BTErase.OnClick     := OnErase;
  end;

  if assigned(GetControl('RBRESTITUE')) then
  begin
    RBRestitue          := TRadioButton(GetControl('RBRESTITUE'));
    RBRestitue.OnClick  := OnRestitue;
  end;

  if assigned(GetControl('RBARECEVOIR')) then
  begin
    RBARestituer         := TRadioButton(GetControl('RBARECEVOIR'));
    RBARestituer.OnClick := OnARecevoir;
  end;

  if assigned(GetControl('RBTOUS')) then
  begin
    RbTous              := TRadioButton(GetControl('RBTOUS'));
    RbTous.OnClick      := OnTous;
  end;

  RBtous.Checked := True;

end ;

procedure TOF_BTFICHEMATCONFIE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTFICHEMATCONFIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFICHEMATCONFIE.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BTFICHEMATCONFIE.OnErase(sender : TObject);
begin

  if Assigned(GetControl('BFF_RESSOURCE'))    then SetControlText('BFF_RESSOURCE',  '');
  if Assigned(GetControl('BFF_RESSOURCE_'))   then SetControlText('BFF_RESSOURCE_', '');
  if Assigned(GetControl('XX_WHERE'))         then XX_WHERE.Text := '';
  if Assigned(GetControl('RBRESTITUE'))       then RBRestitue.Checked   := False;
  if Assigned(GetControl('RBARESTITUER'))     then RBARestituer.Checked := False;
  if Assigned(GetControl('RBTOUS'))           then RBTous.Checked       := False;
        
end;

procedure TOF_BTFICHEMATCONFIE.OnRestitue(sender : TObject);
begin

  XX_WHERE.text := 'BFF_DATEFIN <> "' + USDATETIME(idate2099) + '"';

end;

procedure TOF_BTFICHEMATCONFIE.OnARecevoir(sender : TObject);
begin

  XX_WHERE.text := 'BFF_DATEFIN = "' + USDATETIME(idate2099) + '"';

end;

procedure TOF_BTFICHEMATCONFIE.OnTous(sender : TObject);
begin

  XX_WHERE.text := '';

end;


Initialization
  registerclasses ( [ TOF_BTFICHEMATCONFIE ] ) ;
end.

