{***********UNITE*************************************************
Auteur  ...... : MD
Créé le ...... : 08/07/2004
Modifié le ... :   /  /
Description .. : TOF ancêtre des fiches mul de messagerie
Mots clefs ... : TOF;YMESSAGES
*****************************************************************}
Unit UTofYYMessagesParent ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
{$IFDEF EAGLCLIENT}
     maineagl,
     UTob,
     eMul,
     MenuOLX,
{$ELSE}
     FE_Main,
     Mul,
     MenuOlg,
{$ENDIF}
     HDB,
     HQry,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     AGLInit,
     HDTLinks,
     ExtCtrls,
     Grids, Menus,
     UMailBox;

Type
  TOF_YYMESSAGESPARENT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  public
    procedure mnOuvrirClick (Sender : TObject); overload;
    procedure mnInsertClick (Sender : TObject); overload;
    procedure mnTransfererClick (Sender : TObject); overload;
    procedure mnRepondreClick (Sender : TObject); overload;
    procedure mnSupprimerClick(Sender: TObject); overload;
    procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState); overload;
  private
  protected
    Qry : THQuery;
    Lst : THDBGrid;
    iMessagerieEnCours : Integer;
  end ;

const
  cMsgBoiteReception = 0;
  cMsgBoiteDenvoi    = 1;
  cMsgBrouillons     = 2;

Implementation

uses
    UtilMessages,

    {$IFDEF VER150}
    Variants,
    {$ENDIF}

    galOutil, UtilMulTraitmt, YMessage, YNewMessage, AnnOutils;

//----------------------------------------------------
//--- Nom   : OnArgument
//----------------------------------------------------
procedure TOF_YYMESSAGESPARENT.OnArgument (S : String ) ;
begin
  Inherited ;
  Qry := TFMul(Ecran).Q;
  Lst := TFMul(Ecran).FListe;

  Ecran.OnKeyDown := Form_OnKeyDown;
  if (GetControl ('mnOuvrir')<>nil) then TMenuItem(GetControl('mnOuvrir')).OnClick := mnOuvrirClick;
  if (GetControl ('mnInsert')<>nil) then TMenuItem(GetControl('mnInsert')).OnClick := mnInsertClick;
  if (GetControl ('mnTransferer')<>nil) then TMenuItem(GetControl('mnTransferer')).OnClick := mnTransfererClick;
  if (GetControl ('mnRepondre')<>nil) then TMenuItem(GetControl('mnRepondre')).OnClick := mnRepondreClick;
  if (GetControl ('mnImprimer')<>nil) then TMenuItem(GetControl('mnImprimer')).Visible := False;

  // $$$ JP 14/12/06: suppression possible si concept ok
  if (GetControl ('mnSupprimer')<>nil) then
  begin
       if JaiLeDroitConceptBureau (ccSupprimerMessage) = FALSE then
           SetControlVisible ('mnSupprimer', FALSE)
       else
           TMenuItem(GetControl('mnSupprimer')).OnClick := mnSupprimerClick;
  end;

  // $$$ JP 14/12/06: suppression possible si concept ok
  if GetControl('BDELETE') <> nil then
  begin
         if JaiLeDroitConceptBureau (ccSupprimerMessage) = FALSE then
           SetControlVisible ('BDELETE', FALSE)
       else
           TToolBarButton97(GetControl('BDELETE')).OnClick := mnSupprimerClick;
  end;
end ;

//--------------------------
//--- Nom : mnInsertClick
//--------------------------
procedure TOF_YYMESSAGESPARENT.mnInsertClick(Sender: TObject);
begin
 ShowNewMessage('');
 AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

//--------------------------
//--- Nom : mnOuvrirClick
//--------------------------
procedure TOF_YYMESSAGESPARENT.mnOuvrirClick(Sender: TObject);
begin
  Lst.OnDblClick(Nil);
end;

//----------------------------
//--- Nom : mnRepondreClick
//----------------------------
procedure TOF_YYMESSAGESPARENT.mnRepondreClick(Sender: TObject);
var SMsgGUID : String;
begin
 if Not Lst.Focused then exit;

 if VarIsNull(GetField('YMS_MSGGUID')) then exit;
 //--- On ne répond qu'au message en cours dans la liste
 SmsgGUID := GetField('YMS_MSGGUID');
 if SmsgGUID<>'' then ShowNewMessage('', SmsgGUID, 'RE:');
 //--- Rafraichissement
 TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

//----------------------------
//--- Nom : mnTransfererClick
//----------------------------
procedure TOF_YYMESSAGESPARENT.mnTransfererClick(Sender: TObject);
var SmsgGUID : String;
begin
  if Not Lst.Focused then exit;

  if VarIsNull(GetField('YMS_MSGGUID')) then exit;
  //--- On ne transfère que le message en cours dans la liste
  SmsgGUID := GetField('YMS_MSGGUID');
  if SmsgGUID<>'' then
    ShowNewMessage('', SmsgGUID, 'TR:');

  //--- Rafraichissement
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

//------------------------------
//--- Nom : MnSupprimerClick
//------------------------------
procedure TOF_YYMESSAGESPARENT.mnSupprimerClick(Sender: TObject);
var i : Integer;

  procedure SupprimeMsgDansListe;
  var sUserMail : String;
      SMsgGUID : String;
      iMailId : Integer;
      supprok : Boolean;
  begin
    sUserMail := '';
    iMailId := Qry.FindField('YMS_MAILID').AsInteger;

    SMsgGUID := Qry.FindField('YMS_MSGGUID').AsString;
    if SMsgGUID<>'' then sUserMail := Qry.FindField('YMS_USERMAIL').AsString;

    // Delete table principale
    supprok := True;
    if SMsgGUID<>'' then
      supprok := DeleteMessage(SMsgGUID);

    // Delete YMAILS et dépendances
    if (supprok) and (iMailId<>0) then
      // on ne doit pas passer V_PGI.User car parfois BAL globale
      UserMailDelete(iMailId, sUserMail);
  end;

begin
  // sélection
  if (Lst.NbSelected=0) and (not Lst.AllSelected) then
    begin
    PGIInfo('Aucun message sélectionné.', TitreHalley);
    exit;
    end;
  if PgiAsk('Confirmez-vous la suppression des messages sélectionnés ?',
    'Suppression des messages')=mrNo then exit;
  if Lst.AllSelected then
    begin
{$IFDEF EAGLCLIENT}
    if not TFMul(Ecran).Fetchlestous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
      begin
      Qry.First;
      While Not Qry.Eof do
        begin
        SupprimeMsgDansListe;
        Qry.Next;
        end;
      end;
    end
  else
    begin
    for i:=0 to Lst.NbSelected-1 do
      BEGIN
      Lst.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Qry.TQ.Seek(Lst.Row-1);
{$ENDIF}
      SupprimeMsgDansListe;
      END;
    end;
  // déselectionne
  FinTraitmtMul(TFMul(Ecran));
  // rafraichissement suite aux suppressions
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

//---------------------------
//--- Nom : Form_OnKeyDown
//---------------------------
procedure TOF_YYMESSAGESPARENT.Form_OnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
  //--- F5 = sélection de dossier = double-click sur la grille
  VK_F5 : if (Shift = []) and (Lst.Focused) then
            begin
            Key := 0;
            Lst.OnDblClick(Nil);
            end;

  VK_F9 : TToolbarButton97(GetControl('BCHERCHE')).Click;

  //--- Affichage du menu contextuel
  VK_F11 :  begin
             if GetControl('PPM')<>Nil then
               TPopupMenu(GetControl('PPM')).Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
             //-- évite de traiter le F11 par défaut de Pgi (=vk_choixmul=double clic ds grid)
             Key:=0;
            end;

  //--- Ctrl + Suppr : supprimer
  VK_DELETE : if (Shift = [ssCtrl]) and GetControl('BDELETE').Visible then
               TToolbarButton97(GetControl('BDELETE')).Click;

  //--- CRTL + A
  65 : if (Shift =[ssCtrl]) then
        Lst.AllSelected:=not Lst.AllSelected;

  //--- Ctrl + N : nouveau message
  78 : if (Shift = [ssCtrl]) and GetControl('BINSERT').Visible then
         TToolbarButton97(GetControl('BINSERT')).Click;
  else
    TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
  end;
end;

Initialization
  registerclasses ( [ TOF_YYMESSAGESPARENT ] ) ;
end.

