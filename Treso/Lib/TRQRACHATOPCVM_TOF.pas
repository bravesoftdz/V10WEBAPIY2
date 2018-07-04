{ Unité : Source TOF de la FICHE : TRQRACHATOPCVM
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.20.001.001  09/02/05  JP   Création de l'unité
 7.05.001.001  06/10/06  JP   Gestion des ParamSoc multi sociétés
 7.05.001.001  24/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRQRACHATOPCVM_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  QRS1, FE_Main,
  {$ELSE}
  eQRS1, MaineAGL,
  {$ENDIF}
  Sysutils, HCtrls, UTOF, uAncetreEtat;

type
  TOF_TRQRACHATOPCVM = class (TRANCETREETAT)
    procedure OnArgument(S : string); override;
  private
    TypeEtat : Char;
    ClickImprimer : TNotifyEvent;

    procedure OnAfterShow;
    procedure MajImpression(Sender : TObject);
  end;

procedure TRLanceFiche_EtatsAchatsOPCVM(Nature, Etat : string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  Constantes, UProcGen, HEnt1, Printers, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_EtatsAchatsOPCVM(Nature, Etat : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('TR', 'TRQRACHATOPCVM', '', '', Nature + ';' + Etat + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRACHATOPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  NatutreE : string;
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  AvecParamSoc := True;
  inherited;

  Ecran.HelpContext := 150;

  {Récupération de l'état}
  NatutreE := ReadTokenSt(S);
  TypeEtat := StrToChr(ReadTokenSt(S));

  {Gestion du nom du filtre}
  case TypeEtat of
    tye_TicketOpe : TFQRS1(Ecran).FNomFiltre := 'OVM_ACHAT_TICKET';
    tye_OrdrePaie : TFQRS1(Ecran).FNomFiltre := 'OVM_ACHAT_ORDRE';
    tye_LettreCon : TFQRS1(Ecran).FNomFiltre := 'OVM_ACHAT_LETTRE';
  end;

  {Gestion du paramétrage des états}
  TFQRS1(Ecran).NatureEtat := NatutreE;
  TFQRS1(Ecran).CodeEtat   := NatutreE;
  TFQRS1(Ecran).OnAfterFormShow := OnAfterShow;

  {Gestion des cases à Cocher}
  SetControlVisible('TOP_TICKET'       , TypeEtat = tye_TicketOpe);
  SetControlVisible('TOP_ORDREPAIE'    , TypeEtat = tye_OrdrePaie);
  SetControlVisible('TOP_LETTRECONFIRM', TypeEtat = tye_LettreCon);

  {Gestion de l'impression et de la Maj}
  ClickImprimer := TFQRS1(Ecran).BImprimerClick;
  TFQRS1(Ecran).BImprimer.OnClick := MajImpression;
  {24/10/06 : On filtre les comptes en fonction des sociétés du regroupement Tréso}
  THValComboBox(GetControl('TOP_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TOP_GENERAL')).DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRACHATOPCVM.OnAfterShow;
{---------------------------------------------------------------------------------------}
begin
  {Gestion du caption de l'écran}
  Ecran.Caption := TFQRS1(Ecran).FEtat.Text;
  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRACHATOPCVM.MajImpression(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Suffixe : string;
  Nb : Integer;
begin
  {On force le nombre de copies à 0 ...}
  Printer.Copies := 0;

  {Exécution du code de l'ancêtre}
  if Assigned(ClickImprimer) then ClickImprimer(Sender);

  {Recherche du nom du champ à mettre à jour}
  case TypeEtat of
    tye_TicketOpe : Suffixe := 'TICKET';
    tye_OrdrePaie : Suffixe := 'ORDREPAIE';
    tye_LettreCon : Suffixe := 'LETTRECONFIRM';
  end;

  {... s'il y a eu abandon, le nombre de copies est resté à 0}
  Nb := Printer.Copies;

  if Nb > 0 then
    {Mise à jour du champ précisant que l'impression a été exécutée}
    ExecuteSQL('UPDATE TROPCVM SET TOP_' + Suffixe + ' = "X" WHERE ' + TFQRS1(Ecran).WhereSQL);
end;

initialization
  RegisterClasses([TOF_TRQRACHATOPCVM]);

end.
