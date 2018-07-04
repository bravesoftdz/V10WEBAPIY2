{ Unité : Source TOF de la FICHE : TRQRJALOPCVM : Journal des ventes d'OPCVM
--------------------------------------------------------------------------------------
    Version   |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.2x.xxx.xxx  04/11/04  JP   Création de l'unité
 7.05.001.001  24/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRQRJALOPCVM_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eQRS1, MaineAGL,
  {$ELSE}
  QRS1, FE_Main,
  {$ENDIF}
  SysUtils, HCtrls, UTOF, uAncetreEtat;

type
  TOF_TRQRJALOPCVM = class (TRANCETREETAT)
    procedure OnArgument(S : string); override;
  private
    ckPortefeuille : TCheckBox;
    ckOPCVM        : TCheckBox;
    ckGeneral      : TCheckBox;
    ckEntete       : TCheckBox;

    procedure CasesChange (Sender : TObject);
    procedure PortefChange(Sender : TObject);
  end ;

procedure TRLanceFiche_JalVenteOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  HEnt1, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_JalVenteOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALOPCVM.OnArgument(S : string );
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 150;
  {Gestion des ruptures de l'état}
  ckPortefeuille := TCheckBox(GetControl('PORTEFEUILLE'));
  ckOPCVM        := TCheckBox(GetControl('OPCVM'));
  ckGeneral      := TCheckBox(GetControl('GENERAL'));
  ckEntete       := TCheckBox(GetControl('ENTETE'));
  {Mise à jour de la case Entete}
  ckPortefeuille.OnClick := CasesChange;
  ckOPCVM       .OnClick := CasesChange;
  ckGeneral     .OnClick := CasesChange;
  CasesChange(ckGeneral);
  {Filtre des OPCVM en fonction du portefeuille}
  THValComboBox(GetControl('TVE_PORTEFEUILLE')).OnChange := PortefChange;

  {24/10/06 : On filtre les comptes en fonction des sociétés du regroupement Tréso}
  THValComboBox(GetControl('TVE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TVE_GENERAL')).DataType, '', '');
end;


{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALOPCVM.CasesChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {La case Entete est invisible. Je l'ai rajoutée car la condition ci-dessous s'avérait avoir
   un erratique dans les ruptures ou dans les formules de l'état. Il s'agit d'imprimer l'entête
   et le pied de groupe Général si l'on demande une rupture par Compte ou aucune rupture}
  ckEntete.Checked := ckGeneral.Checked or
                      (not ckPortefeuille.Checked and not ckGeneral.Checked and not ckOPCVM.Checked);

  {Définition des champs sur lesquels s'appliquent les ruptures de l'état en fonction
   des options sélectionnées, sachant que la rupture 3 est toujours imprimée car elle
   contient les entêtes de colonnes => on part donc du niveau le plus bas puis on remonte :
    Général, Opcvm, Portefeuille}
  if ckEntete.Checked then begin
    SetControlText('XX_RUPTURE3', 'TVE_GENERAL');
    if ckOPCVM.Checked then begin
      SetControlText('XX_RUPTURE2', 'TVE_CODEOPCVM');
      if ckPortefeuille.Checked then
        SetControlText('XX_RUPTURE1', 'TVE_PORTEFEUILLE');
    end
    else if ckPortefeuille.Checked then
      SetControlText('XX_RUPTURE2', 'TVE_PORTEFEUILLE');
  end
  else if ckOPCVM.Checked then begin
    SetControlText('XX_RUPTURE3', 'TVE_CODEOPCVM');
    if ckPortefeuille.Checked then
      SetControlText('XX_RUPTURE2', 'TVE_PORTEFEUILLE');
  end
  else
    SetControlText('XX_RUPTURE3', 'TVE_PORTEFEUILLE');

end;

{Gestion du filtre sur les portefeuilles
{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALOPCVM.PortefChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ch : string;
begin
  Ch := GetControlText('TVE_PORTEFEUILLE');
  if Ch <> '' then
    THMultiValComboBox(GetControl('TVE_CODEOPCVM')).Plus := 'TOF_PORTEFEUILLE = "' + Ch + '"'
  else
    THMultiValComboBox(GetControl('TVE_CODEOPCVM')).Plus := '';
end;

initialization
  RegisterClasses([TOF_TRQRJALOPCVM]);

end.

