{ Unité : Source TOF de la FICHE : TRQRVENTEOPCVM
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.20.001.001  09/02/05  JP   Création de l'unité
 7.05.001.001  06/10/06  JP   Gestion des ParamSoc multi sociétés
 7.05.001.001  24/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRQRVENTEOPCVM_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eQRS1,
  {$ELSE}
  FE_Main, QRS1,
  {$ENDIF}
  SysUtils, HCtrls, UTOF, uAncetreEtat;

type
  TOF_TRQRVENTEOPCVM = class (TRANCETREETAT)
    procedure OnArgument(S : string); override;
  private
    TypeEtat : Char;
    ClickImprimer : TNotifyEvent;

    procedure OnAfterShow;
    procedure MajImpression(Sender : TObject);
  end;

procedure TRLanceFiche_EtatsVenteOPCVM(Nature, Etat : string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  Constantes, UProcGen, HEnt1, Printers, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_EtatsVenteOPCVM(Nature, Etat : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('TR', 'TRQRVENTEOPCVM', '', '', Nature + ';' + Etat + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRVENTEOPCVM.OnArgument(S : string);
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
    tye_TicketOpe : TFQRS1(Ecran).FNomFiltre := 'OVM_VENTE_TICKET';
    tye_OrdrePaie : TFQRS1(Ecran).FNomFiltre := 'OVM_VENTE_ORDRE';
    tye_LettreCon : TFQRS1(Ecran).FNomFiltre := 'OVM_VENTE_LETTRE';
  end;

  {Gestion du paramétrage des états}
  TFQRS1(Ecran).NatureEtat := NatutreE;
  TFQRS1(Ecran).CodeEtat   := NatutreE;
  TFQRS1(Ecran).OnAfterFormShow := OnAfterShow;

  {Gestion des cases à Cocher}
  SetControlVisible('TVE_TICKET'       , TypeEtat = tye_TicketOpe);
  SetControlVisible('TVE_ORDREPAIE'    , TypeEtat = tye_OrdrePaie);
  SetControlVisible('TVE_LETTRECONFIRM', TypeEtat = tye_LettreCon);

  {Gestion de l'impression et de la Maj}
  ClickImprimer := TFQRS1(Ecran).BImprimerClick;
  TFQRS1(Ecran).BImprimer.OnClick := MajImpression;

  {24/10/06 : On filtre les comptes en fonction des sociétés du regroupement Tréso}
  THValComboBox(GetControl('TVE_GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TVE_GENERAL')).DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRVENTEOPCVM.OnAfterShow;
{---------------------------------------------------------------------------------------}
begin
  Ecran.Caption := TFQRS1(Ecran).FEtat.Text;
  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRVENTEOPCVM.MajImpression(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Suffixe : string;
  Nb : Integer;
begin
  {On force le nombre de copies à 0 ...}
  Printer.Copies := 0;

  {Exécution du code de l'ancêtre}
  if Assigned(ClickImprimer) then ClickImprimer(Sender);

  case TypeEtat of
    tye_TicketOpe : Suffixe := 'TICKET';
    tye_OrdrePaie : Suffixe := 'ORDREPAIE';
    tye_LettreCon : Suffixe := 'LETTRECONFIRM';
  end;

  {... s'il y a eu abandon, le nombre de copies est resté à 0}
  Nb := Printer.Copies;

  if Nb > 0 then
    {Mise à jour du champ précisant que l'impression a été exécutée}
    ExecuteSQL('UPDATE TRVENTEOPCVM SET TVE_' + Suffixe + ' = "X" WHERE ' + TFQRS1(Ecran).WhereSQL);
end;

initialization
  RegisterClasses([TOF_TRQRVENTEOPCVM]);

end.



