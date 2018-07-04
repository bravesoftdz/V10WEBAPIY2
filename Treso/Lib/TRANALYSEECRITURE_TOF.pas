{Source de la tof TRANALYSEECRITURE
--------------------------------------------------------------------------------------
  Version     |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
  0.91         26/11/03  JP   création de l'unité
07.05.001.001  10/10/06  JP   FQ 10370 : Oracle ne supporte pas les jointures externes sur 2 tables
07.05.001.001  23/10/06  JP   Gestion du multi sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.001.010  20/09/07  JP   FQ 10370 suite : on modifie pour toutes les version d'oracle et on vide TFStat(Ecran).FSQL
--------------------------------------------------------------------------------------}
unit TRANALYSEECRITURE_TOF;

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
  TOF_TRANALYSEECRITURE = class (TOFCONF)
  {$ELSE}
  TOF_TRANALYSEECRITURE = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
  protected
    DeviseAff : string;
    MultiSoc  : Boolean; {14/11/06 : Si depuis le menu analyses multi sociétés}
    procedure DeviseOnChange (Sender : TObject);
    procedure NoDossierChange(Sender : TObject);
  end ;

function TRLanceFiche_TRANALYSEECRITURE(Dom, Fiche, Range, Lequel, Arguments : string) : string;

Implementation

uses
   Commun, ExtCtrls, Constantes, Stat, UtilPGI;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRANALYSEECRITURE(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  {!!!! Problème sur le retour en eAgl}
  Result := AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRANALYSEECRITURE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 150;
  MultiSoc := not (Ecran.Name = 'TRANALYSEECRITURE');
  THValComboBox(GetControl('TE_DEVISE')).OnChange := DeviseOnChange;
  {10/10/06 : FQ 10370 : adaptation sous Oracle
   20/09/07 : FQ 10370 suite : on modifie pour toutes les version d'oracle et on vide TFStat(Ecran).FSQL}
  if isOracle then begin
    TFStat(Ecran).FSQL.Lines.Clear;
    TFStat(Ecran).FSQL.Lines.Add('SELECT * FROM TRECRITURE , BANQUECP, CIB ');
    TFStat(Ecran).FSQL.Lines.Add('WHERE BQ_CODE = TE_GENERAL AND TCI_CODECIB = TE_CODECIB AND TCI_BANQUE = BQ_BANQUE');
  end;
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
procedure TOF_TRANALYSEECRITURE.NoDossierChange(Sender : TObject);
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
procedure TOF_TRANALYSEECRITURE.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DeviseAff := GetControlText('TE_DEVISE');
  if DeviseAff = '' then DeviseAff := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), DeviseAff);
end;

initialization
  RegisterClasses([TOF_TRANALYSEECRITURE]);

end.

