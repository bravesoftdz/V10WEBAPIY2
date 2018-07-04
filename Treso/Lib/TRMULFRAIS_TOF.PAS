{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.xx.xxx.xxx    10/07/04   JP   Création de l'unité

--------------------------------------------------------------------------------------}
unit TRMULFRAIS_TOF;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, HCtrls, HEnt1, HMsgBox, UTOF;

type
  TOF_TRMULFRAIS = class (TOF)
    procedure OnArgument(S : string); override;
  private
    procedure BInsertClick(Sender : TObject);
    procedure BDeleteClick(Sender : TObject);
    procedure ListDblClick(Sender : TObject);
  end ;

procedure TRLanceFiche_MulFrais(Dom, Fiche, Range, Lequel, Arguments : string);


implementation

uses TomFrais, AglInit, HTB97;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulFrais(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFRAIS.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  TToolbarButton97(GetControl('BINSERT')).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BDeleteClick;
  TFMul(Ecran).FListe.OnDblClick := ListDblClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFRAIS.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Frais('TR', 'TRFRAIS', '', '', ActionToString(taCreatEnSerie));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFRAIS.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if TFMul(Ecran).FListe.nbSelected = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '');
    Exit;
  end;
  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer la ligne en cours ?;Q;YNC;N;C;', '', '') = mrYes then begin
    for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
      TFMul(Ecran).FListe.GotoLeBookmark(n);
      ExecuteSQL('DELETE FROM FRAIS WHERE TFR_CONTRAT = "' + GetField('TFR_CONTRAT')  +
                 '" AND TFR_CODEFRAIS = "'+ GetField('TFR_CODEFRAIS') + '"');
    end;
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULFRAIS.ListDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Frais('TR', 'TRFRAIS', '', GetField('TFR_CONTRAT') + ';' + GetField('TFR_CODEFRAIS'), ActionToString(taModif));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;


initialization
  RegisterClasses([TOF_TRMULFRAIS]);


end.
