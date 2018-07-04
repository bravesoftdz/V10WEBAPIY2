{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.2X.xxx.xxx    03/11/04   JP   Création de l'unité  : Mul de saisie des cours des OPCVM
--------------------------------------------------------------------------------------}
unit TRCOTATIONOPCVM_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL, UTOB,
  {$ELSE}
  FE_Main, Mul, DBGrids, db, HDB,
  {$ENDIF}
  Classes, HTB97, Menus, UTOF, Controls, SysUtils;

type
  TOF_TRCOTATIONOPCVM = class(TOF)
    procedure OnArgument(S : string); override;
  private
    BOuvrir   : TToolBarButton97;
    BInsert   : TToolBarButton97;
    BDelete   : TToolBarButton97;
    BFirst    : TToolBarButton97;
    BCherche  : TToolBarButton97;
    PopupMenu : TPopUpMenu;

    procedure BOuvrirOnClick(Sender : TObject);
    procedure BInsertOnClick(Sender : TObject);
    procedure BDeleteOnClick(Sender : TObject);
    procedure DeviseChange  (Sender : TObject);
  end;

procedure TRLanceFiche_TRCOTATIONOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  HEnt1, HMsgbox, AglInit, HCtrls, TRCOTATIONOPCVM_TOM, Commun, ExtCtrls{TImage};

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_TRCOTATIONOPCVM(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCOTATIONOPCVM.BOuvrirOnClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  str    : string;
  strD   : string;
  phrase : string;
begin
  strD := VarToStr(GetField('TTO_DATE'));
  str  := VarToStr(GetField('TTO_CODEOPCVM'));
  phrase := str + ';' + strD;
  TRLanceFiche_TRCoursOPCVM('TR', 'TRCOURSOPCVM', '', phrase, 'ACTION=MODIFICATION;MUL;' + phrase);
  if Bcherche <> nil then BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCOTATIONOPCVM.BInsertOnClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  c : string;
begin
  c := GetControlText('TTO_CODEOPCVM') + ';';
  TRLanceFiche_TRCoursOPCVM('TR', 'TRCOURSOPCVM', '',  '', ActionToString(taCreat) + ';MUL;' + c);
  if Bcherche <> nil then BCherche.Click;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCOTATIONOPCVM.BDeleteOnClick(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  p  : Integer;
  ch : string;

    {----------------------------------------------------------------------}
    procedure SupprimerEnreg;
    {----------------------------------------------------------------------}
    var
      s : string;
      d : string;
    begin
      s := TFMul(Ecran).Q.FindField('TTO_CODEOPCVM').AsString;
      d := UsDateTime(TFMul(Ecran).Q.FindField('TTO_DATE').AsDateTime);
      Inc(p);
      ExecuteSQL('DELETE FROM TRCOTATIONOPCVM WHERE TTO_CODEOPCVM = "' + s + '" AND TTO_DATE = "' + d + '"');
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

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer le(s) cours sélectionné(s)' +
                  ' ?;Q;YNC;N;C;', '', '') = mrYes then begin
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
  end
  else
    ch := TraduireMemoire('Aucun enregistrement n''a été supprimé.');
  PGIInfo(ch);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCOTATIONOPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  inherited;
  ch := ReadTokenSt(S);
  if ch <> '' then
    SetControlText('TTO_CODEOPCVM', ch);
  Ecran.HelpContext := 150;
  BFirst   := TToolBarButton97(GetControl('BFIRST'));
  bOuvrir  := TToolBarButton97(GetControl('BOUVRIR'));
  BDelete  := TToolBarButton97(GetControl('BDELETE'));
  BInsert  := TToolBarButton97(GetControl('BINSERT'));
  BCherche := TToolBarButton97(GetControl('BCHERCHE'));
  {$IFDEF EAGLCLIENT}
  SetControlVisible('BSELECTALL', False);
  {$ELSE}
  SetControlVisible('BSELECTALL', True);
  {$ENDIF}

  if BOuvrir <> nil then BOuvrir.OnClick   := BOuvrirOnClick;
  if BFirst  <> nil then BFirst.Visible    := False;
  if BInsert <> nil then BInsert.OnClick   := BInsertOnClick;
  if BDelete <> nil then BDelete.OnClick   := BDeleteOnClick;

  TFMul(Ecran).FListe.OnDblClick := BOuvrirOnClick;

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.Items[0].OnClick := BInsertOnClick;
  PopupMenu.Items[1].OnClick := BDeleteOnClick;
  AddMenuPop(PopupMenu, '', '');
  {Pour le chargement de la devise}
  THValComboBox(GetControl('TTO_CODEOPCVM')).OnChange := DeviseChange;
  DeviseChange(GetControl('TTO_CODEOPCVM'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRCOTATIONOPCVM.DeviseChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Devise : string;
begin
  Devise := GetControlText('TTO_CODEOPCVM');
  Devise := RetDeviseOPCVM(Devise);
  AssignDrapeau(TImage(GetControl('IDEV')), Devise);
  SetControlCaption('DEV', Devise);
end;

initialization
  RegisterClasses([TOF_TRCOTATIONOPCVM]);

end.
      
