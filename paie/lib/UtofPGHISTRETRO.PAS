{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTRETRO ()
Mots clefs ... : TOF;PGHISTRETRO
*****************************************************************}
Unit UtofPGHISTRETRO ;

Interface

Uses
{$IFDEF VER150}
     Variants,
{$ENDIF}
     StdCtrls,
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
  TOF_PGHISTRETRO = Class (TOF)
    public

    CombNaturerub : THValComboBox;

    procedure ChangeNature ( Sender : Tobject );

    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_PGHISTRETRO.OnArgument (Arguments : String ) ;
var
  LeWhere : string;
  Naturerub : String;
  Rubrique : String;
  Datevalidite : String;
begin
  Inherited ;
  if Arguments<>'' then
  begin
    NatureRub := READTOKENST(Arguments);
    Rubrique := READTOKENST(Arguments);
    Datevalidite := READTOKENST(Arguments);

    LeWhere := 'PGT_NATURERUB="'+Naturerub+'" AND PGT_RUBRIQUE="'+Rubrique+
                '" AND PGT_DATEVALIDITE="'+Datevalidite+'"';
    SetControlText('XX_WHERE', LeWhere);
  end;

  CombNaturerub := THValComboBox(GetControl('PGT_NATURERUB'));
  CombNaturerub.OnChange := ChangeNature;

end ;

procedure TOF_PGHISTRETRO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETRO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETRO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETRO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETRO.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETRO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETRO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGHISTRETRO.ChangeNature(Sender: Tobject);
var
  Rubrique : THEdit;
  Trubrique : THLabel;

begin
  Rubrique := THEdit(GetControl('PGT_RUBRIQUE'));
  Trubrique := THLabel(GetControl('TPGT_RUBRIQUE'));
  SetControlText('PGT_RUBRIQUE','');

  if GetControlText('PGT_NATURERUB')='' then
  begin
    Rubrique.ElipsisButton := False;
    Rubrique.DataType := '';
    Trubrique.Caption := 'Rubrique';
  end;

  if GetControlText('PGT_NATURERUB')='COT' then
  begin
    Rubrique.DataType := 'PGCOTIS';
    Rubrique.ElipsisButton := True;
    Trubrique.Caption := 'Cotisation';
  end;

  if GetControlText('PGT_NATURERUB')='AAA' then
  begin
    Rubrique.DataType := 'PGREMUNERATION';
    Rubrique.ElipsisButton := True;
    Trubrique.Caption := 'Rémunération';
  end;

  if GetControlText('PGT_NATURERUB')='BAS' then
  begin
    Rubrique.ElipsisButton := False;
    Rubrique.DataType := '';
    Trubrique.Caption := 'Rubrique';
    SetControlText('PGT_NATURERUB','');
  end;

end;

Initialization
  registerclasses ( [ TOF_PGHISTRETRO ] ) ;
end.
