{Source de la tof TRCUBETRANSACT
--------------------------------------------------------------------------------------
  Version     |  Date  | Qui  |   Commentaires
--------------------------------------------------------------------------------------
  0.91         26/11/03  JP   création de l'unité
07.05.001.001  23/10/06  JP   Gestion du multi sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}

unit TRCUBETRANSACT_TOF;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  StdCtrls, Controls, Classes, HCtrls, HEnt1;

type
  {$IFDEF TRCONF}
  TOF_TRCUBETRANSACT = class (TOFCONF)
  {$ELSE}
  TOF_TRCUBETRANSACT = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument (S : string ) ; override ;
  protected
    DeviseAff : string;
    MultiSoc  : Boolean; {14/11/06 : Si depuis le menu analyses multi sociétés}
    procedure DeviseOnChange (Sender: TObject);
    procedure NoDossierChange(Sender : TObject);
  end;

function TRLanceFiche_TRCUBETRANSACT(Dom, Fiche, Range, Lequel, Arguments : string) : string;

implementation

uses
   Commun, ExtCtrls, Constantes;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRCUBETRANSACT(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  {!!!! Problème sur le retour en eAgl}
  Result := AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCUBETRANSACT.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 150;
  THValComboBox(GetControl('BQ_DEVISE')).OnChange := DeviseOnChange;
  MultiSoc := not (Ecran.Name = 'TRCUBETRANSACT');
  {23/10/06 : gestion du multi sociétés}
  SetControlVisible('TCT_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTCT_NODOSSIER', IsTresoMultiSoc);

  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  THValComboBox(GetControl('TCT_COMPTETR')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TCT_COMPTETR')).DataType, '', '');
  if not MultiSoc then begin
    THMultiValComboBox(GetControl('TCT_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    THMultiValComboBox(GetControl('TCT_NODOSSIER')).OnChange := NoDossierChange;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCUBETRANSACT.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := GetControlText('TCT_NODOSSIER');
  if THMultiValComboBox(GetControl('TCT_NODOSSIER')).Tous then s := '';
  THValComboBox(GetControl('TCT_COMPTETR')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TCT_COMPTETR')).DataType, '', s);
  THValComboBox(GetControl('TCT_COMPTETR')).ItemIndex := - 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCUBETRANSACT.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DeviseAff := GetControlText('BQ_DEVISE');
  if DeviseAff = '' then DeviseAff := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), DeviseAff);
end;

initialization
  RegisterClasses([TOF_TRCUBETRANSACT]);

end.

