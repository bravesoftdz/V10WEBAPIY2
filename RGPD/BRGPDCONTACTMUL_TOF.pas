{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
Unit BRGPDCONTACTMUL_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , FE_Main
  {$ENDIF EAGLCLIENT}
  , uTob
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HEnt1
  , HMsgBox
  , UTOF
  , CBPMcd
  , BRGPDMUL_TOF
  , BRGPDUtils
  ;

function BLanceFiche_RGPDContactMul(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BRGPDCONTACTMUL = Class (TOF_BRGPDMUL)
  private
    ContactAdr : THCheckbox;

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

uses
  utilPGI
  , wCommuns
  , mul
  ;

function BLanceFiche_RGPDContactMul(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

procedure TOF_BRGPDCONTACTMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDCONTACTMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDCONTACTMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDCONTACTMUL.OnLoad ;
begin
  Inherited ;
  if (ContactAdr.Checked) and (IsConsent) then
    SetControlText('XX_WHERE', ' AND EXISTS (SELECT 1 FROM ADRESSES WHERE ADR_REFCODE = C_TIERS AND ADR_NUMEROCONTACT = C_NUMEROCONTACT)')
  else
    SetControlText('XX_WHERE', '');
end ;

procedure TOF_BRGPDCONTACTMUL.OnArgument (S : String ) ;
begin
  sPopulationCode    := RGPDContact;
  sFieldCode         := 'C_TIERS';
  sFieldCode2        := 'C_NUMEROCONTACT';
  sFieldCode3        := '';
  sFieldLabel        := 'C_NOM';
  sFieldLabel2nd     := 'C_PRENOM';
  Inherited ;
  ContactAdr         := THCheckbox(GetControl('CONTACTADR'));
  ContactAdr.Visible := (RgpdAction = rgdpaConsentRequest);
  if RgPDAction = rgpdaAnonymization then
    THCheckBox(GetControl('C_PRINCIPAL')).State := cbUnchecked
  else
    THCheckBox(GetControl('C_PRINCIPAL')).State := cbGrayed;
end ;

procedure TOF_BRGPDCONTACTMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDCONTACTMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDCONTACTMUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BRGPDCONTACTMUL ] ) ;
end.

