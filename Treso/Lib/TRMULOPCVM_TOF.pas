{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.0X.xxx.xxx    02/11/04   JP   Création de l'unité
--------------------------------------------------------------------------------------}
unit TRMULOPCVM_TOF;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL, uTob, 
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, UTOF, Menus;

type
  TOF_TRMULOPCVM = class (TOF)
    procedure OnArgument(S : string); override;
  private
    PopupMenu : TPopUpMenu;
    procedure GererBouton;
    procedure GererPopup;
  protected

    procedure BInsertClick   (Sender : TObject);
    procedure BDeleteClick   (Sender : TObject);
    procedure ListDblClick   (Sender : TObject);
  end ;

procedure TRLanceFiche_MulOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);


implementation

uses
  TROPCVMREF_TOM, AglInit, HTB97;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULOPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  {Affectation des évènements aux Boutons}
  GererBouton;
  {Affectation des évènements aux menus}
  GererPopup;
  {$IFDEF EAGLCLIENT}
  SetControlVisible('BSELECTALL', False);
  {$ELSE}
  SetControlVisible('BSELECTALL', True);
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULOPCVM.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_OPCVMREF('TR', 'TROPCVMREF', '', '', ActionToString(taCreatEnSerie));

  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULOPCVM.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  p  : Integer;
  Ok : Boolean;
  ch : string;

    {----------------------------------------------------------------------}
    procedure SupprimerEnreg;
    {----------------------------------------------------------------------}
    var
      s : string;
    begin
      s := TFMul(Ecran).Q.FindField('TOF_CODEOPCVM').AsString;
      if not ExisteSQL('SELECT TOP_CODEOPCVM FROM TROPCVM WHERE TOP_CODEOPCVM = "' + s + '"') then begin
        Inc(p);
        ExecuteSQL('DELETE FROM TROPCVMREF WHERE TOF_CODEOPCVM = "' + s + '"');
      end
      else
        Ok := True;
    end;

begin
  {Aucune sélection, on sort}
  if (TFMul(Ecran).FListe.NbSelected = 0)
  {$IFNDEF EAGLCLIENT}
  and not TFMul(Ecran).FListe.AllSelected
  {$ENDIF}
  then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '');
    Exit;
  end;

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer le(s) OPCVM sélectionné(s)' +
                  ' ?;Q;YNC;N;C;', '', '') = mrYes then begin
    Ok := False;
    p  := 0;
    BeginTrans;
    try
      {$IFNDEF EAGLCLIENT}
      TFMul(Ecran).Q.First;
      if TFMul(Ecran).FListe.AllSelected then
        while not TFMul(Ecran).Q.EOF do begin
          SupprimerEnreg;
          TFMul(Ecran).Q.Next;
        end
      else
      {$ENDIF}

        for n := 0 to TFMul(Ecran).FListe.nbSelected - 1 do begin
          TFMul(Ecran).FListe.GotoLeBookmark(n);
          {$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
          {$ENDIF}
          SupprimerEnreg;
        end;

      CommitTrans;
    except
      on E : Exception do begin
        RollBack;
        PGIError(E.Message);
        Exit;
      end;
    end;
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
  if p > 0 then begin
    if p = 1 then
      ch := TraduireMemoire('Un enregistrement a été supprimé.')
    else
      ch := TraduireMemoire(IntToStr(p) + ' enregistrements ont été supprimés.');
    if Ok then
      ch := ch + #13 + TraduireMemoire('Cependant certains enregistrements n''ont pu être supprimés car ils sont utilisés en gestion financière.');
  end
  else
    ch := TraduireMemoire('Aucun enregistrement n''a été supprimé, car les OPCVM sont utilisés en gestion financière.');
  PGIInfo(ch);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULOPCVM.ListDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_OPCVMREF('TR', 'TROPCVMREF', '', GetField('TOF_CODEOPCVM'), ActionToString(taModif));
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULOPCVM.GererBouton;
{---------------------------------------------------------------------------------------}
begin
  TToolbarButton97(GetControl('BINSERT')).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BDeleteClick;
  TToolbarButton97(GetControl('BOUVRIR')).OnClick := ListDblClick;
  TFMul(Ecran).FListe.OnDblClick := ListDblClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULOPCVM.GererPopup;
{---------------------------------------------------------------------------------------}
begin
  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.Items[0].OnClick := BInsertClick;
  PopupMenu.Items[1].OnClick := BDeleteClick;
  AddMenuPop(PopupMenu, '', '');
end;


initialization
  RegisterClasses([TOF_TRMULOPCVM]);

end.

