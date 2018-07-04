{***********UNITE*************************************************
Auteur  ...... : O. TARCY
Créé le ...... : 31/08/2000
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : ECARTCAISSE
Suite ........ : Gestion des écarts dans le contrôle de la caisse
Mots clefs ... : TOF;UTOFMFOECARTCAISSE;FO
*****************************************************************}
unit UtofMFOEcartCaisse;

interface
uses
  Classes, Forms, SysUtils, UTOF, Vierge, HCtrls;

type
  TOF_MFOECARTCAISSE = class(TOF)
  public
    procedure OnArgument(stargument: string); override;
  end;

implementation

procedure TOF_MFOECARTCAISSE.OnArgument(stargument: string);
var ligneforce: THLabel;
  F: TForm;
  x: integer;
  Critere, ChampCritere, ValeurCritere: string;
begin
  inherited;
  F := TForm(Ecran);

  repeat
    Critere := Trim(ReadTokenSt(stargument));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampCritere := copy(Critere, 1, x - 1);
        ValeurCritere := copy(Critere, x + 1, length(Critere));
        if (ChampCritere = 'ONGLET') and (ValeurCritere <> 'ESPECES') then
        begin
          SetControlVisible('BMODIFIER', False);
          ligneforce := THLabel(TFVierge(F).FindComponent('TEXTE3'));
          ligneforce.caption := '';
        end;
      end;
    end;
  until Critere = '';
end;

initialization
  registerclasses([TOF_MFOECARTCAISSE]);
end.
