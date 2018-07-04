{ Unité : Source TOF de la FICHE : TRQRVENTEOPCVM
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
               30/01/02  BT   Création de l'unité
 6.20.001.001  09/02/05  JP   Mise en conformité avec les OPCVM
 7.05.001.001  23/10/06  JP   Gestion des filtres multi sociétés
                              Mise en place de l'ancêtre des états
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TofLanceEtat ;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eQRS1 ,
  {$ELSE}
  QRS1 , FE_Main,
  {$ENDIF}
  SysUtils, HCtrls, UTOF, uAncetreEtat;


type
  TOF_LANCEETAT = class(TRANCETREETAT)
    procedure OnArgument(S : string); override;
  protected
    Devise	: THValComboBox;
    DeviseAff	: string;
    Nature	: string ;
    Modele	: string ;
    TypeEtat    : Char;
    ClickImprimer : TNotifyEvent;

    procedure OnAfterShow;
  public
    {On surcharge, car le champ n'est pas "_DEVISE"}
    procedure DeviseOnChange (Sender : TObject); override;
    procedure MajImpression  (Sender : TObject);
    procedure NoDossierChange(Sender : TObject);
  end;

procedure TRLanceFiche_EtatsCourtsTermes(Nature, Etat : string);


implementation


uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  UProcGen, HEnt1, Constantes, Commun, Printers, ExtCtrls;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_EtatsCourtsTermes(Nature, Etat : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('TR', 'TRLANCEETAT', '', '', Nature + ';' + Etat + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEETAT.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DeviseAff := GetControlText('TCT_DEVMONTANT');
  SetControlCaption('TCT_DEVMONTANT', DeviseAff);
  AssignDrapeau(TImage(GetControl('IDEV')), DeviseAff);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEETAT.OnArgument(S : string );
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
    tye_TicketOpe : TFQRS1(Ecran).FNomFiltre := 'CT_TICKET';
    tye_OrdrePaie : TFQRS1(Ecran).FNomFiltre := 'CT_ORDRE';
    tye_LettreCon : TFQRS1(Ecran).FNomFiltre := 'CT_LETTRE';
  end;

  {Gestion du paramétrage des états}
  TFQRS1(Ecran).NatureEtat := NatutreE;
  TFQRS1(Ecran).CodeEtat   := NatutreE;
  TFQRS1(Ecran).OnAfterFormShow := OnAfterShow;

  {Gestion des cases à Cocher}
  SetControlVisible('TCT_TICKET'       , TypeEtat = tye_TicketOpe);
  SetControlVisible('TCT_ORDREPAIE'    , TypeEtat = tye_OrdrePaie);
  SetControlVisible('TCT_LETTRECONFIRM', TypeEtat = tye_LettreCon);

  {Gestion de l'impression et de la Maj}
  ClickImprimer := TFQRS1(Ecran).BImprimerClick;
  TFQRS1(Ecran).BImprimer.OnClick := MajImpression;

  {08/08/06 : gestion du multi sociétés}
  SetControlVisible('TCT_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTCT_NODOSSIER', IsTresoMultiSoc);

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  if not EtatMD then begin
    THValComboBox(GetControl('TCT_COMPTETR')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TCT_COMPTETR')).DataType, '', '');
    THValComboBox(GetControl('TCT_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    THValComboBox(GetControl('TCT_NODOSSIER')).OnChange := NoDossierChange;
  end;

  {Gestion des devises}
  THValComboBox(GetControl('TCT_DEVMONTANT')).OnChange := DeviseOnChange;
  DeviseAff := GetControlText('TCT_DEVMONTANT');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEETAT.OnAfterShow;
{---------------------------------------------------------------------------------------}
begin
  Ecran.Caption := TFQRS1(Ecran).FEtat.Text;
  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEETAT.MajImpression(Sender : TObject);
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
    ExecuteSQL('UPDATE COURTSTERMES SET TCT_' + Suffixe + ' = "X" WHERE ' + TFQRS1(Ecran).WhereSQL);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_LANCEETAT.NoDossierChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  THValComboBox(GetControl('TCT_COMPTETR')).Plus := FiltreBanqueCp(THValComboBox(GetControl('TCT_COMPTETR')).DataType, '', GetControlText('TCT_NODOSSIER'));
  SetControlText('TCT_COMPTETR', '');
end;

initialization
  RegisterClasses([TOF_LANCEETAT]);

end.

