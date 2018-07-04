{ Unité : Source TOF de la FICHE : TRDECALAGEJOUR
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            11/09/03    JP     Création de l'unité
 6.01.001.001    17/11/04    JP     FQ 10181 bis : le décalage en jours ne se faisait que pour
                                    un nombre positif : on ne pouvait revenir en arrière 
--------------------------------------------------------------------------------------}
unit TRDECALAGEJOUR_TOF ;

interface

uses
  StdCtrls,  Controls,  Classes,
  {$IFNDEF EAGLCLIENT}
  FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  SysUtils, HCtrls, UTOF, Vierge;

type
  TOF_TRDECALAGEJOUR = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate                 ; override ;
  end ;

function TRLanceFiche_DecalageJours(var DateValeur : string; var NbJours : Integer) : Boolean;


implementation

var
  Valeur : string;
  Jours  : Integer;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_DecalageJours(var DateValeur : string; var NbJours : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Valeur := DateTostr(StrToDate(DateValeur) + 1);
  Jours  := NbJours;
  Result := AGLLanceFiche('TR', 'TRDECALAGEJOUR', '', '', '') <> '';
  DateValeur := Valeur;
  NbJours    := Jours;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDECALAGEJOUR.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  Ecran.HelpContext := 150;
  SetControlText('TE_DATEVALEUR', Valeur);
  {17/11/04 FQ 10181 : on gère maintenant les valeurs négatives}
  THSpinEdit(GetControl('SP_JOURS')).MinValue := - 99;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDECALAGEJOUR.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  {On demande d'incrémenter les dates de valeurs de x jours}
  if TRadioButton(GetControl('RB_JOURS')).Checked then
    Jours  := THSpinEdit(GetControl('SP_JOURS')).Value
  {On demande de réaffecter les dates de valeur}
  else
    Valeur := GetControlText('TE_DATEVALEUR');
  inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDECALAGEJOUR.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  TFVierge(Ecran).Retour := 'VALIDER';
end;

initialization
  RegisterClasses ( [ TOF_TRDECALAGEJOUR ] ) ;

end.
