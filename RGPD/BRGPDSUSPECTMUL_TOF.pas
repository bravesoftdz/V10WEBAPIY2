{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
Unit BRGPDSUSPECTMUL_TOF ;

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

function BLanceFiche_RGPDSuspectMul(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BRGPDSUSPECTMUL = Class (TOF_BRGPDMUL)
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

function BLanceFiche_RGPDSuspectMul(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

procedure TOF_BRGPDSUSPECTMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDSUSPECTMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDSUSPECTMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDSUSPECTMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDSUSPECTMUL.OnArgument (S : String ) ;
begin
  sPopulationCode := RGPDSuspect;
  sFieldCode      := 'RSU_SUSPECT';
  sFieldCode2     := '';
  sFieldCode3     := '';  
  sFieldLabel     := 'RSU_LIBELLE';
  sFieldLabel2nd  := 'RSU_PRENOM';
  Inherited ;
end ;

procedure TOF_BRGPDSUSPECTMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDSUSPECTMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDSUSPECTMUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BRGPDSUSPECTMUL ] ) ;
end.

