{ Unité : Source TOF de la FICHE : TRQRRELEVEBQ
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91          17/11/03  JP   Création de l'unité
 7.05.001.001  24/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRQRRELEVEBQ_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  QRS1, FE_Main,
  {$ELSE}
  eQRS1, MaineAGL,
  {$ENDIF}
  SysUtils, UTOF, uAncetreEtat;

type
  TOF_TRQRRELEVEBQ = class (TRANCETREETAT)
    procedure OnArgument(S : string); override;
    procedure OnLoad                ; override;
  private
  end;

procedure TRLanceFiche_Releve(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  HCtrls, Commun, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Releve(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRRELEVEBQ.OnArgument(S : string );
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 150;
  {24/10/06 : On filtre les comptes en fonction des sociétés du regroupement Tréso}
  if not EtatMD then
    THValComboBox(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TE_GENERAL')).DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRRELEVEBQ.OnLoad;
{---------------------------------------------------------------------------------------}
var
  Deb : string;
begin
  inherited;
  Deb := DateToStr(DebutAnnee(StrToDate(GetControlText('TE_DATEVALEUR'))));
  {Mise à jour des critères de début et de fin qui permettent de simuler un extrait bancaire}
  SetControlText('DATEDEB', GetControlText('TE_DATEVALEUR'));
  SetControlText('DATEFIN', GetControlText('TE_DATEVALEUR_'));
  SetControlText('DEBANNEE', Deb);
end;

initialization
  RegisterClasses([TOF_TRQRRELEVEBQ]);
                       
end.
