{Source de la tof TRCUBEECRITURE
--------------------------------------------------------------------------------------
  Version     | Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
  0.91         26/11/03  JP   création de l'unité
07.05.001.001  10/10/06  JP   FQ 10370 : Oracle ne supporte pas les jointures externes sur 2 tables
07.05.001.001  23/10/06  JP   Gestion du multi sociétés
08.00.001.025  16/07/07  JP   FQ 10370 : suite : cela ne marche pas non plus sous Oracle 10.
                              Par ailleurs, manifestement, l'agl a changé ...
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}

unit TRCUBEECRITURE_TOF;

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
  StdCtrls, Controls, Classes, HCtrls, HEnt1, Cube;

type
  {$IFDEF TRCONF}
  TOF_TRCUBEECRITURE = class (TOFCONF)
  {$ELSE}
  TOF_TRCUBEECRITURE = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
    procedure OnUpdate              ; override;
  protected
    DeviseAff : string;
    MultiSoc  : Boolean; {14/11/06 : Si depuis le menu analyses multi sociétés}
    procedure DeviseOnChange (Sender : TObject);
    procedure NoDossierChange(Sender : TObject);
  end ;

function TRLanceFiche_TRCUBEECRITURE(Dom, Fiche, Range, Lequel, Arguments : string) : string;

implementation

uses
   Commun, ExtCtrls, UTilPgi;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRCUBEECRITURE(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  {!!!! Problème sur le retour en eAgl}
  Result := AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCUBEECRITURE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';' ;
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 150;
  MultiSoc := not (Ecran.Name = 'TRCUBEECRITURE');
  THValComboBox(GetControl('TE_DEVISE')).OnChange := DeviseOnChange;
  {10/10/06 : FQ 10370 : adaptation sous Oracle
   16/07/07 : FQ 10370 suite : manifestement, cela ne marche pas mieux sous Oracle 10 => on généralise
              Par ailleurs, le fonctionnement de l'agl semble avoir changé}
  if isOracle {(V_PGI.Driver in [dbORACLE7, dbORACLE8, dbORACLE9])} then
    //TFCube(Ecran).FromSQL := 'FROM TRECRITURE, BANQUECP, CIB WHERE BQ_CODE = TE_GENERAL AND TCI_CODECIB = TE_CODECIB AND TCI_BANQUE = BQ_BANQUE';
    TFCube(Ecran).FromSQL := 'TRECRITURE, BANQUECP, CIB ';

  {23/10/06 : gestion du multi sociétés}
  SetControlVisible('TE_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTE_NODOSSIER', IsTresoMultiSoc);

  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  THValComboBox(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TE_GENERAL')).DataType, '', '');
  if not MultiSoc then begin
    THMultiValComboBox(GetControl('TE_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    THMultiValComboBox(GetControl('TE_NODOSSIER')).OnChange := NoDossierChange;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCUBEECRITURE.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := GetControlText('TE_NODOSSIER');
  if THMultiValComboBox(GetControl('TE_NODOSSIER')).Tous then s := '';
  THValComboBox(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TE_GENERAL')).DataType, '', s);
  THValComboBox(GetControl('TE_GENERAL')).ItemIndex := - 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCUBEECRITURE.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DeviseAff := GetControlText('TE_DEVISE');
  if DeviseAff = '' then DeviseAff := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), DeviseAff);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCUBEECRITURE.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {16/07/07 : le fonctionnement de l'agl semble avoir changé}
  if isOracle then begin
    if TFCube(Ecran).WhereSQL <> '' then
      TFCube(Ecran).WhereSQL := 'AND BQ_CODE = TE_GENERAL AND TCI_CODECIB = TE_CODECIB AND TCI_BANQUE = BQ_BANQUE '
    else
      TFCube(Ecran).WhereSQL := 'BQ_CODE = TE_GENERAL AND TCI_CODECIB = TE_CODECIB AND TCI_BANQUE = BQ_BANQUE '
  end;
end;

initialization
  RegisterClasses([TOF_TRCUBEECRITURE]);

end.

