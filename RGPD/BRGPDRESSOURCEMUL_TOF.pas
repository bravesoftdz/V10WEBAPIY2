{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
Unit BRGPDRESSOURCEMUL_TOF ;

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

function BLanceFiche_RGPDResourceMul(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BRGPDRESSOURCEMUL = Class (TOF_BRGPDMUL)
  private

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
   BRGPDVALIDTRT_TOF
  , FormsName
  ;

function BLanceFiche_RGPDResourceMul(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

procedure TOF_BRGPDRESSOURCEMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDRESSOURCEMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDRESSOURCEMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDRESSOURCEMUL.OnLoad ;
begin
  Inherited ;
  SetcontrolText('XX_WHERE', '(ARS_TYPERESSOURCE IN("SAL", "INT"))');
end ;

procedure TOF_BRGPDRESSOURCEMUL.OnArgument (S : String ) ;
begin
  sPopulationCode := RGPDResource;
  sFieldCode      := 'ARS_RESSOURCE';
  sFieldCode2     := '';
  sFieldCode3     := '';
  sFieldLabel     := 'ARS_LIBELLE';
  sFieldLabel2nd  := '';
  Inherited ;
  THValCombobox(GetControl('ARS_TYPERESSOURCE')).Plus := ' AND (CC_CODE IN("SAL", "INT"))';
end ;

procedure TOF_BRGPDRESSOURCEMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDRESSOURCEMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDRESSOURCEMUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BRGPDRESSOURCEMUL ] ) ;
end.

