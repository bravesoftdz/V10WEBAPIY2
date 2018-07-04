{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
Unit BRGPDTIERSMUL_TOF ;

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

function BLanceFiche_RGPDThirdMul(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BRGPDTIERSMUL = Class (TOF_BRGPDMUL)
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

function BLanceFiche_RGPDThirdMul(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

procedure TOF_BRGPDTIERSMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDTIERSMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDTIERSMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDTIERSMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDTIERSMUL.OnArgument (S : String ) ;
begin
  sPopulationCode := RGPDThird;
  sFieldCode      := 'T_TIERS';
  sFieldCode2     := 'T_AUXILIAIRE';
  sFieldCode3     := '';
  sFieldLabel     := 'T_LIBELLE';
  sFieldLabel2nd  := 'T_PRENOM'; 
  Inherited ;
end ;

procedure TOF_BRGPDTIERSMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDTIERSMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDTIERSMUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BRGPDTIERSMUL ] ) ;
end.

