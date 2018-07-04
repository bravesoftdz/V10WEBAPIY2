{***********UNITE*************************************************
Auteur  ...... : MD
Créé le ...... : 08/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYREPERTPERSO ()
Mots clefs ... : TOF;YYREPERTPERSO
*****************************************************************}
Unit UTofYYRepertPerso ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     menus,
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     MaineAGL,
{$ELSE}
     db,
     dbtables,
     mul,
     Fe_Main,
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
     AGLInit,
     YNewMessage,
     Mailol,
     windows ;

Type
  TOF_YYREPERTPERSO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  public
    procedure BINSERT_OnClick(Sender : TObject);
    procedure BDELETE_OnClick(Sender : TObject);
    procedure FLISTE_OnDblClick(Sender : TObject);
    procedure BIMPORTCARNET_OnClick(Sender : TObject);
  private
    PMRepertPerso    : TPopupMenu;
    MnNouveauMessage : TMenuItem;

    {$IFDEF BUREAU}
    Tel1Menu     :TMenuItem;
    Tel2Menu     :TMenuItem;

    procedure    OnAppelPopup (Sender:TObject);
    procedure    OnTel1Click  (Sender:TObject);
    procedure    OnTel2Click  (Sender:TObject);
    {$ENDIF}

    procedure mnNouveauMessageClick(Sender: TObject);
    procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;


//////////// IMPLEMENTATION ////////////
Implementation

uses AssistImportCarnet,
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     UtilMulTraitmt
{$IFDEF BUREAU}
     ,entDP, galMenuDisp
{$ENDIF}
     ;

procedure TOF_YYREPERTPERSO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYREPERTPERSO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYREPERTPERSO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYREPERTPERSO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_YYREPERTPERSO.OnArgument (S : String ) ;
var
    vTel     : Variant;
    strTel   : String;
begin
 Inherited ;
 // Répertoire personnel : on ne consulte que ses propres contacts...
 SetControlText('XX_WHERE', 'YRP_USER="'+V_PGI.User+'"');

 //--- Création du menu PopUp
 PMRepertPerso := TPopupMenu.Create (Ecran);
 PMRepertPerso.Images := FMenuDisp.SmallImages;

 //--- Création des Items du menu PopUp
 MnNouveauMessage:=TMenuItem.Create (Ecran);
 if (MnNouveauMessage<>nil) then
  begin
   MnNouveauMessage.Caption := '&Nouveau message';
   MnNouveauMessage.OnClick := mnNouveauMessageClick;
   PMRepertPerso.Items.Add (MnNouveauMessage);
  end;

 {$IFDEF BUREAU}
 if VH_DP.ctiAlerte <> nil then
  begin
   Tel1Menu := TMenuItem.Create (Ecran);
   if Tel1Menu <> nil then
    begin
     strTel := '';
     vTel   := GetField ('YRP_TEL1');
     if (VarIsEmpty (vTel) = FALSE) and (VarIsNull (vTel) = FALSE) then
        strTel := Trim (VarAsType (vTel, varString));
     if strTel <> '' then
     begin
      Tel1Menu.Caption := 'Appeler le ' + strTel;
      Tel1Menu.Hint    := strTel;
      Tel1Menu.Visible := TRUE;
     end
     else
     begin
      Tel1Menu.Caption := '';
      Tel1Menu.Hint    := '';
      Tel1Menu.Visible := FALSE;
     end;

     Tel1Menu.ImageIndex := cImgMenuDispTel;
     Tel1Menu.OnClick    := OnTel1Click;
     PMRepertPerso.Items.Add (Tel1Menu);
    end;

   Tel2Menu := TMenuItem.Create (Ecran);
   if Tel2Menu <> nil then
    begin
     strTel := '';
     vTel   := GetField ('YRP_TEL2');
     if (VarIsEmpty (vTel) = FALSE) and (VarIsNull (vTel) = FALSE) then
      strTel := Trim (VarAsType (vTel, varString));
     if strTel <> '' then
      begin
       Tel2Menu.Caption := 'Appeler le ' + strTel;
       Tel2Menu.Hint    := strTel;
       Tel2Menu.Visible := TRUE;
      end
     else
      begin
       Tel2Menu.Caption := '';
       Tel2Menu.Hint    := '';
       Tel2Menu.Visible := FALSE;
      end;

     Tel2Menu.ImageIndex := cImgMenuDispTel;
     Tel2Menu.OnClick    := OnTel2Click;
     PMRepertPerso.Items.Add (Tel2Menu);
    end;
  end;
{$ENDIF}

 //--- Gestion des evenements
 Ecran.OnKeyDown:=Form_OnKeyDown;
 THDBGrid (GetControl ('FLISTE')).PopupMenu := PMRepertPerso;
 THDBGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;

 TButton(GetControl('BINSERT')).OnClick := BINSERT_OnClick;
 TButton(GetControl('BDELETE')).OnClick := BDELETE_OnClick;
 if GetControl('BIMPORTCARNET')<>Nil then TButton(GetControl('BIMPORTCARNET')).OnClick := BIMPORTCARNET_OnClick;

end ;

procedure TOF_YYREPERTPERSO.OnClose;
begin
 inherited;
 {$IFDEF BUREAU}
  FreeAndNil (Tel1Menu);
  FreeAndNil (Tel2Menu);
 {$ENDIF}
 FreeAndNil (PmRepertPerso);
end;

procedure TOF_YYREPERTPERSO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYREPERTPERSO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_YYREPERTPERSO.BINSERT_OnClick(Sender: TObject);
begin
  AglLanceFiche('YY', 'YYREPERTPERSO', '', '', 'ACTION=CREATION');
  AglRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

procedure TOF_YYREPERTPERSO.BDELETE_OnClick(Sender: TObject);
var i : Integer;
    GuidRep : String;
    Qry : THQuery;
    Lst : THDBGrid;
begin
  Lst := THDBGrid(GetControl('FLISTE'));
  Qry := TFMul(Ecran).Q;
  // sélection
  if (Lst.NbSelected=0) and (not Lst.AllSelected) then
    begin
    PGIInfo('Aucun élément sélectionné', TitreHalley);
    exit;
    end;
  if PGIAsk('Vous allez supprimer la sélection.#13#10 Confirmez-vous ?', TitreHalley)=mrYes then
    begin
    if Lst.AllSelected then
      BEGIN
{$IFDEF EAGLCLIENT}
      if not TFMul(Ecran).Fetchlestous then
        PGIInfo('Impossible de récupérer tous les enregistrements')
      else
{$ENDIF}
        begin
        while Not Qry.EOF do
          BEGIN
          GuidRep := Qry.FindField('YRP_GUIDREP').AsString;
          ExecuteSQL('DELETE FROM YREPERTPERSO WHERE YRP_GUIDREP="'+GuidRep+'"'
            // facultatif, car clé GuidRep unique, indépt des users (voir UTomYRepertPerso)
            + ' AND YRP_USER="'+V_PGI.User+'"');
          Qry.Next;
          END;
        end;
      END
    ELSE
      BEGIN
      for i:=0 to Lst.NbSelected-1 do
        BEGIN
        Lst.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
        Qry.TQ.Seek(Lst.Row - 1) ;
{$ENDIF}
        GuidRep := Qry.FindField('YRP_GUIDREP').AsString;
        ExecuteSQL('DELETE FROM YREPERTPERSO WHERE YRP_GUIDREP="'+GuidRep+'"'
          + ' AND YRP_USER="'+V_PGI.User+'"');
        END;
      END;
    end;
  // déselectionne
  FinTraitmtMul(TFMul(Ecran));
  AglRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

procedure TOF_YYREPERTPERSO.FLISTE_OnDblClick(Sender: TObject);
begin
     if VarIsNull (GetField ('YRP_GUIDREP')) then
        exit;

     // si on veut scroller sur les enreg
{$IFDEF EAGLCLIENT}
     TheMulQ:=TFMul(Ecran).Q.TQ;
{$ELSE}
     TheMulQ:=TFMul(Ecran).Q;
{$ENDIF}

     // Ouverture fiche répertoire
     AglLanceFiche ('YY', 'YYREPERTPERSO', '', GetField ('YRP_GUIDREP') + ';' + GetField ('YRP_USER'), 'ACTION=MODIFICATION');

     // Mise à jour de la liste suite aux modif' éventuelles
     AglRefreshDB ([LongInt(Ecran), 'FListe'], 2);
end;

procedure TOF_YYREPERTPERSO.BIMPORTCARNET_OnClick(Sender: TObject);
begin
  LanceAssistImportCarnet;
  AglRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{$IFDEF BUREAU}
procedure TOF_YYREPERTPERSO.OnAppelPopup (Sender:TObject);
var
   vTel     :variant;
   strTel   :string;
begin
end;

procedure TOF_YYREPERTPERSO.OnTel1Click (Sender:TObject);
begin
     if PgiAsk ('Appeler le ' + Tel1Menu.Hint  + ' ?') = mrYes then
        VH_DP.ctiAlerte.MakeCall (Tel1Menu.Hint);
end;

procedure TOF_YYREPERTPERSO.OnTel2Click (Sender:TObject);
begin
     if PgiAsk ('Appeler le ' + Tel2Menu.Hint  + ' ?') = mrYes then
        VH_DP.ctiAlerte.MakeCall (Tel2Menu.Hint);
end;
{$ENDIF}

//----------------------------------
//--- Nom : MnNouveauMessageClick
//----------------------------------
procedure TOF_YYREPERTPERSO.mnNouveauMessageClick(Sender: TObject);
begin
 if (VarIsNull(GetField ('YRP_EMAIL'))) then Exit;

 //--- Envoit de l'Email
 if (VH_DP.SeriaMessagerie) then
  ShowNewMessage ('','','',GetField ('YRP_EMAIL'))
end;

//---------------------------
//--- Nom : Form_OnKeyDown
//---------------------------
procedure TOF_YYREPERTPERSO.Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 case Key of
  VK_F11 :  begin
             TPopupMenu(GetControl('PM_REPERTPERSO')).Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
             Key := 0;
            end;
 end;
end;

initialization
  registerclasses ( [ TOF_YYREPERTPERSO ] ) ;

end.



