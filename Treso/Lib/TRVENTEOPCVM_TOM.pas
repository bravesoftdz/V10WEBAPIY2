{-------------------------------------------------------------------------------------
    Version  |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
6.20.001.001   21/12/04   JP   Création de l'unité
6.50.001.001   20/05/05   JP   Gestion des frais et commissions
7.09.001.001   23/10/06   JP   Gestion des filtres Multi sociétés
8.01.001.001   26/12/06   JP   Accès à la fiche d'achat des OPCVM
--------------------------------------------------------------------------------------}

unit TRVENTEOPCVM_TOM;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HEnt1, UTOM;

type
  TOM_TRVENTEOPCVM = class (TOM)
    procedure OnArgument(S : string); override;
    procedure OnLoadRecord          ; override;
  private
    {20/05/05 : Certains montants sont stockés uniquement en devise pivot dans la
                base afin d'éviter de dupliquer tous les champs}
    procedure AffichageEnDevises;
    {26/12/06 : Affichage du détail des OPCVM}
    procedure BOpcvmOnClick(Sender : TObject);
  end;

procedure TRLanceFiche_TRVenteOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  Commun, Constantes, HCtrls, ComCtrls, TROPCVM_TOM, HTB97, AglInit;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_TRVenteOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRVENTEOPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('TVE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TVE_GENERAL')).DataType, '', '');
  {$ELSE}
  THDBValComboBox(GetControl('TVE_GENERAL')).Plus := FiltreBanqueCp(THDBValComboBox(GetControl('TVE_GENERAL')).DataType, '', '');
  {$ENDIF EAGLCLIENT}
  (GetControl('PGVENTE') as TPageControl).ActivePageIndex := 0;
  (GetControl('BOPCVM') as TToolbarButton97).OnClick := BOpcvmOnClick; 
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRVENTEOPCVM.OnLoadRecord;
{---------------------------------------------------------------------------------------}
var
  Dev : string;
  n   : Integer;
begin
  {Chargement du drapeau relatif à la devise de la vente en cours}
  Dev := GetField('TVE_DEVISE');
  MajAffichageDevise(GetControl('IDEV'), GetControl('DEV'), Dev, sd_Aucun);
  for n := 1 to 5  do
    MajAffichageDevise(GetControl('IDEV' + IntToStr(n)), GetControl('DEV' + IntToStr(n)), Dev, sd_Aucun);

  {Conversion des montants des frais et commissions en devise et affichage}
  AffichageEnDevises;
end;

{20/05/05 : Certains montants sont stockés uniquement en devise pivot dans la
            base afin d'éviter de dupliquer tous les champs
{---------------------------------------------------------------------------------------}
procedure TOM_TRVENTEOPCVM.AffichageEnDevises;
{---------------------------------------------------------------------------------------}
var
  Mnt : Double;
  Cot : Double;
begin
  {Récupération de la Cotation}
  Cot := Valeur(VarToStr(GetField('TVE_COTATION')));
  {Gestion des frais en devises}
  Mnt := Valeur(VarToStr(GetField('TVE_FRAISEUR')));
  SetControlText('EDFRAIS', FloatToStr(Mnt * Cot));
  {Gestion des frais en Commissions en devises
   21/12/06 : en mettant Mnt := , cela ira mieux}
  Mnt := Valeur(VarToStr(GetField('TVE_COMMISSIONEUR')));
  SetControlText('EDCOM', FloatToStr(Mnt * Cot));
  {Gestion de la Tva en devises sur les frais et les commissions
   21/12/06 : en mettant Mnt := , cela ira mieux}
  Mnt := Valeur(VarToStr(GetField('TVE_TVAVENTEEUR')));
  SetControlText('EDTVA', FloatToStr(Mnt * Cot));
end;

{26/12/06 : Affichage du détail des OPCVM
{---------------------------------------------------------------------------------------}
procedure TOM_TRVENTEOPCVM.BOpcvmOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_OPCVM('TR', 'TROPCVM', '', VarToStr(GetField('TVE_NUMTRANSAC')), ActionToString(taConsult));
end;

initialization
  RegisterClasses([TOM_TRVENTEOPCVM]);

end.
