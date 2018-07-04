{Source de la tof TRANALYSETRANSACT
--------------------------------------------------------------------------------------
  Version     |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
  0.91         25/11/03  JP   création de l'unité
07.05.001.001  23/10/06  JP   Gestion du multi sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}

unit TRANALYSETRANSACT_TOF;

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
  Controls, Classes;

type
  {$IFDEF TRCONF}
  TOF_TRANALYSETRANSACT = class (TOFCONF)
  {$ELSE}
  TOF_TRANALYSETRANSACT = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
  protected
    MultiSoc  : Boolean; {14/11/06 : Si depuis le menu analyses multi sociétés}
    procedure NoDossierChange(Sender : TObject);
  end;

function TRLanceFiche_TRANALYSETRANSACT(Dom, Fiche, Range, Lequel, Arguments : string) : string;

implementation

uses
  Commun, Constantes, HCtrls;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRANALYSETRANSACT(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  {!!!! Problème sur le retour en eAgl}
  Result := AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRANALYSETRANSACT.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 150;
  MultiSoc := not (Ecran.Name = 'TRANALYSETRANSACT');
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
procedure TOF_TRANALYSETRANSACT.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := GetControlText('TCT_NODOSSIER');
  if THMultiValComboBox(GetControl('TCT_NODOSSIER')).Tous then s := '';
  THValComboBox(GetControl('TCT_COMPTETR')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TCT_COMPTETR')).DataType, '', s);
  THValComboBox(GetControl('TCT_COMPTETR')).ItemIndex := - 1;
end;



initialization
  RegisterClasses([TOF_TRANALYSETRANSACT]);

end.

