{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
Unit BRGPDUTILISATMUL_TOF ;

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

function BLanceFiche_RGPDUtilisatMul(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BRGPDUTILISATMUL = Class (TOF_BRGPDMUL)
  private
    UtilAdr : THCheckbox;

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
  , wCommuns
  , UtilPGI
  , Mul
  ;

function BLanceFiche_RGPDUtilisatMul(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

procedure TOF_BRGPDUTILISATMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDUTILISATMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDUTILISATMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDUTILISATMUL.OnLoad ;
begin
  Inherited ;
  if (UtilAdr.Checked) and (IsConsent) then
    SetControlText('XX_WHERE', ' AND EXISTS (SELECT 1 FROM SALARIES WHERE PSA_SALARIE = US_AUXILIAIRE AND CONCAT(PSA_ADRESSE1, PSA_ADRESSE2, PSA_ADRESSE3, PSA_CODEPOSTAL, PSA_VILLE) <> "")')
  else
    SetControlText('XX_WHERE', '');
end ;

procedure TOF_BRGPDUTILISATMUL.OnArgument (S : String ) ;
begin
  sPopulationCode := RGPDUser;
  sFieldCode      := 'US_UTILISATEUR';
  sFieldCode2     := 'US_AUXILIAIRE';
  sFieldCode3     := '';
  sFieldLabel     := 'US_LIBELLE';
  sFieldLabel2nd  := '';
  Inherited ;
  UtilAdr         := THCheckbox(GetControl('UTILADR'));
  UtilAdr.Visible := (RgpdAction = rgdpaConsentRequest);
end ;

procedure TOF_BRGPDUTILISATMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDUTILISATMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDUTILISATMUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BRGPDUTILISATMUL ] ) ;
end.

