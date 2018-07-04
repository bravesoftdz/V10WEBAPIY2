{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.xx.xxx.xxx    15/07/04   JP   Création de l'unité
 6.0X.xxx.xxx    04/08/04   JP   Gestion des PopupMenus
--------------------------------------------------------------------------------------}
unit TRMULCONTRAT_TOF;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, UTOF, Menus;

type
  TOF_TRMULCONTRAT = class (TOF)
    procedure OnArgument(S : string); override;
  private
    PopupMenu : TPopUpMenu; {04/08/04}

    procedure BInsertClick   (Sender : TObject);
    procedure BDeleteClick   (Sender : TObject);
    procedure ListDblClick   (Sender : TObject);
    procedure BDupliquerClick(Sender : TObject);
  end ;

procedure TRLanceFiche_MulContrat(Dom, Fiche, Range, Lequel, Arguments : string);


implementation

uses
  TRCONTRAT_TOM, AglInit, HTB97;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulContrat(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONTRAT.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  TToolbarButton97(GetControl('BINSERT'   )).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BDELETE'   )).OnClick := BDeleteClick;
  TToolbarButton97(GetControl('BOUVRIR'   )).OnClick := ListDblClick;
  TToolbarButton97(GetControl('BDUPLIQUER')).OnClick := BDupliquerClick;
  TFMul(Ecran).FListe.OnDblClick := ListDblClick;

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU')); {04/08/04}
  PopupMenu.Items[0].OnClick := BInsertClick;
  PopupMenu.Items[1].OnClick := BDupliquerClick;
  PopupMenu.Items[2].OnClick := BDeleteClick;
  AddMenuPop(PopupMenu, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONTRAT.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Contrat('TR', 'TRCONTRAT', '', '', ActionToString(taCreatEnSerie));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONTRAT.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
begin
  if TFMul(Ecran).FListe.nbSelected = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '');
    Exit;
  end;

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer le(s) contrat(s) sélectionné(s)'#13 +
                  'et tous les frais rattachés ?;Q;YNC;N;C;', '', '') = mrYes then begin
    {On boucle sur la sélection}
    for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
      TFMul(Ecran).FListe.GotoLeBookmark(n);
      BeginTrans;
      try
        {Suppression du contrat courant ...}
        ExecuteSQL('DELETE FROM TRCONTRAT WHERE TRC_CODECONTRAT = "' + GetField('TRC_CODECONTRAT') + '"');
        {... et des frais qui lui sont rattachés}
        ExecuteSQL('DELETE FROM FRAIS WHERE TFR_CONTRAT = "' + GetField('TRC_CODECONTRAT') + '"');
        CommitTrans;
      except
        on E : Exception do begin
          RollBack;
          PGIError(E.Message);
        end;
      end;
    end;
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONTRAT.ListDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Contrat('TR', 'TRCONTRAT', '', GetField('TRC_CODECONTRAT'), ActionToString(taModif));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{Création d'un nouveau contrat et duplication de tous les frais atenants
{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCONTRAT.BDupliquerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Contrat('TR', 'TRCONTRAT', '', '', ActionToString(taCreat) + ';DUPLI;' + GetField('TRC_CODECONTRAT'));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

initialization
  RegisterClasses([TOF_TRMULCONTRAT]);

end.

