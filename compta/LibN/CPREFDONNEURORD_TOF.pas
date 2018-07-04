unit CPREFDONNEURORD_TOF;
{-------------------------------------------------------------------------------------
  Version     |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
06.50.001.001  09/05/05   JP    Création de l'unité : FQ 15513
06.50.001.005  17/06/05   JP    Gestion du VCOM 1000 : la longueur du retour n'est pas
                                la même (12 au lieu de 7)
08.00.001.018  14/05/07   JP    FQ 18135 : La longueur pour GEFACTOFRANCE est 17
--------------------------------------------------------------------------------------}
interface

uses
    StdCtrls, Controls, Classes, UTof,
    {$IFDEF EAGLCLIENT}
    MaineAGL,
    {$ELSE}
    FE_Main,
    {$ENDIF}
    Vierge, Forms, SysUtils, ComCtrls, HTB97;

type
    TOF_CPREFDONNEURORD = class(TOF)
      procedure OnUpdate; override;
      procedure OnArgument(S : string); override;
      procedure OnClose; override;
    private
      {JP 17/06/05 : Longueur de la chaine de retour}
      Longueur : Byte;
      type_ecran : string;
      function  IsChecked : Boolean;
      {JP 17/06/05 : On précise sur quel format on travaille : 400 ou 1000}
      procedure cb1000OnClick(Sender : TObject);
      procedure BFermeClick(Sender : TObject);
    public
      property VCOM1000 : Boolean read IsChecked;
    end;

function CPRefDonneurOrd(cex_Type : String) : string;

const
  cex_1000 = 'VCOM1000';
  cex_VCOM = 'VCO';
  cex_GEF  = 'GEF';

implementation

uses
  HCtrls, HEnt1;



{---------------------------------------------------------------------------------------}
function CPRefDonneurOrd(cex_Type : String) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AGLLanceFiche('CP', 'CPREFDONNEURORD', '', '', cex_Type);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPREFDONNEURORD.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TFVierge(Ecran).Retour := '';
  Longueur := 7;
  TCheckBox(GetControl('CB1000')).OnClick := cb1000OnClick;
  TToolBarButton97(Getcontrol('BFERME'  , True)).OnClick := BFermeClick;

  {YMO 03/04/2006 Dans le cas de GEFACTOFRANCE, on change les libellés
   car ici il ne s'agit plus de donneur d'ordres}
  if S = cex_GEF then begin
    Type_ecran := cex_GEF;
    TFVierge(Ecran).Caption := 'Référence';
    UpdateCaption(Ecran);
    TCheckBox(GetControl('CB1000')).Visible := False;
    {YMO FQ 17817 10/04/2006}
    THLabel(GetControl('LBCAPTION')).Caption := 'Référence de l''export';
    {JP 14/05/07 : FQ 18135 : Si la variable longueur n'est pas renseignée, le MaxLength ne sert à rien}
    Longueur := 17;
    THEdit(GetControl('NUMCFONB')).MaxLength := Longueur;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPREFDONNEURORD.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if VCOM1000 and (TFVierge(Ecran).Retour = '') then
    TFVierge(Ecran).Retour := cex_1000;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPREFDONNEURORD.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TFVierge(Ecran).Retour := Copy(GetControlText('NUMCFONB'), 1, Longueur);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPREFDONNEURORD.cb1000OnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if VCOM1000 then begin
    SetControlCaption('LBCAPTION', TraduireMemoire('Référence du paiement'));
    THEdit(GetControl('NUMCFONB')).MaxLength := 12;
    Longueur := 12;
  end
  else begin
    SetControlCaption('LBCAPTION', TraduireMemoire('Référence du donneur d''ordre'));
    THEdit(GetControl('NUMCFONB')).MaxLength := 7;
    Longueur := 7;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPREFDONNEURORD.IsChecked: Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := GetCheckBoxState('CB1000') = cbChecked;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/06/2006
Modifié le ... :   /  /    
Description .. : Pour GEFACTO, on doit différencier validation avec reference vide
Description .. : et sortie par la croix
Mots clefs ... : FQ18356
*****************************************************************}
procedure TOF_CPREFDONNEURORD.BFermeClick(Sender : TObject);
begin
    If type_ecran=cex_GEF then TFVierge(Ecran).Retour := '@stop';
end;

initialization
  RegisterClasses([TOF_CPREFDONNEURORD]);
    

end.
