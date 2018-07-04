{Source de la tof TRANALYSEVIREMT
--------------------------------------------------------------------------------------
  Version     |  Date  | Qui  |   Commentaires
--------------------------------------------------------------------------------------
  0.91         25/11/03  JP   création de l'unité
 7.05.001.001  23/10/06  JP   Gestion du multi sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}

unit TRANALYSEVIREMENT_TOF;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  {$IFDEF TRCONF}
  UlibConfidentialite,
  {$ENDIF TRCONF}
  StdCtrls, Controls, Classes, HCtrls, HEnt1, UTOF;

type
  {$IFDEF TRCONF}
  TOF_TRANALYSEVIREMENT = class (TOFCONF)
  {$ELSE}
  TOF_TRANALYSEVIREMENT = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
    procedure OnLoad                ; override;
  protected
    DeviseAff : string;
    MultiSoc  : Boolean; {14/11/06 : Si depuis le menu analyses multi sociétés}
    procedure DeviseOnChange  (Sender : TObject);
    procedure DNoDossierChange(Sender : TObject);
    procedure SNoDossierChange(Sender : TObject);
  end;

function TRLanceFiche_TRANALYSEVIREMENT(Dom, Fiche, Range, Lequel, Arguments : string) : string;

Implementation

uses
   Commun, ExtCtrls, Constantes;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRANALYSEVIREMENT(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  {!!!! Problème sur le retour en eAgl}
  Result := AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRANALYSEVIREMENT.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  MultiSoc := not (Ecran.Name = 'TRANALYSEVIREMENT');
  THValComboBox(GetControl('TEQ_DEVISE')).OnChange := DeviseOnChange;
  {23/10/06 : gestion du multi sociétés}
  SetControlVisible('TEQ_DNODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTEQ_DNODOSSIER', IsTresoMultiSoc);
  SetControlVisible('TEQ_DNODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTEQ_DNODOSSIER', IsTresoMultiSoc);

  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  THValComboBox(GetControl('TEQ_DGENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TEQ_DGENERAL')).DataType, '', '');
  THValComboBox(GetControl('TEQ_SGENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TEQ_SGENERAL')).DataType, '', '');
  if not MultiSoc then begin
    THMultiValComboBox(GetControl('TEQ_SNODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    THMultiValComboBox(GetControl('TEQ_SNODOSSIER')).OnChange := SNoDossierChange;
    THMultiValComboBox(GetControl('TEQ_DNODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    THMultiValComboBox(GetControl('TEQ_DNODOSSIER')).OnChange := DNoDossierChange;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRANALYSEVIREMENT.OnLoad;
{---------------------------------------------------------------------------------------}
{$IFDEF TRCONF}
var
  DConf : string;
  SConf : string;
  SQLConf : string;
{$ENDIF TRCONF}
begin
  inherited;
  {$IFDEF TRCONF}
  DConf := '';
  SConf := '';
  SQLConf := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if SQLConf <> '' then SQLConf := ' AND (' + SQLConf + ') ';
  DConf   := AliasSQL(SQLConf, 'BQ', 'BD');
  SConf   := AliasSQL(SQLConf, 'BQ', 'BS');
  XX_WHERECONF.Text := DConf + SConf;
  {$ENDIF TRCONF}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRANALYSEVIREMENT.SNoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := GetControlText('TEQ_SNODOSSIER');
  if THMultiValComboBox(GetControl('TEQ_SNODOSSIER')).Tous then s := '';
  THValComboBox(GetControl('TEQ_SGENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TEQ_SGENERAL')).DataType, '', s);
  THValComboBox(GetControl('TEQ_SGENERAL')).ItemIndex := - 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRANALYSEVIREMENT.DNoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := GetControlText('TEQ_DNODOSSIER');
  if THMultiValComboBox(GetControl('TEQ_DNODOSSIER')).Tous then s := '';
  THValComboBox(GetControl('TEQ_DGENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TEQ_DGENERAL')).DataType, '', s);
  THValComboBox(GetControl('TEQ_DGENERAL')).ItemIndex := - 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRANALYSEVIREMENT.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DeviseAff := GetControlText('TEQ_DEVISE');
  if DeviseAff = '' then DeviseAff := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), DeviseAff);
end;

initialization
  RegisterClasses([TOF_TRANALYSEVIREMENT]);

end.

