{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/06/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYUTILISATRESSMUL ()
Mots clefs ... : TOF;YYUTILISATRESSMUL
*****************************************************************}
unit UTOFYYUTILISATRESSMUL ;

interface

Uses StdCtrls,
     Controls,
     Classes,
     menus,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     mul,
     fe_main,
{$ELSE}
     eMul,
     uTob,
     maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HDB,
     UTOF,
     YNewMessage,
     Mailol,
     windows,
     utofzoneslibres
     ;

Type
  TOF_YYUTILISATRESSMUL = class (TOF_ZONESLIBRES) // class( TOF)

  public
         procedure OnArgument (strArgs:string); override;
         procedure OnClose;                     override;

{$IFDEF BUREAU}
  private
         PMUTILISATRESS   : TPopupMenu;
         MnNouveauMessage : TMenuItem;
         TelMenu          : TMenuItem;
         Tel1Menu         : TMenuItem;
         Tel2Menu         : TMenuItem;
         Tel3Menu         : TMenuItem;
         bModeSelect      : Boolean; // Utilisation de la fiche pour faire une sélection
         bMultiSelect     : boolean; // $$$ JP 02/04/07: multisélection

         procedure OnPopupUtilisatRess (Sender:TObject);
         procedure AppelerClick  (Sender:TObject); // $$$ JP 23/05/07
         procedure FListe_OnDblClick(Sender: TObject);
         procedure mnNouveauMessageClick(Sender: TObject);
         procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{$ENDIF}

         procedure BMultiSelectOuvrirClick (Sender:TObject);
  end;

function YYLanceUtilisatRessMul (strArgs:string) : String;


/////////// IMPLEMENTATION ////////////
implementation

{$IFDEF BUREAU}
uses
    entDP, galMenuDisp,
    dpOutils,
    htb97, hqry // $$$ JP 02/04/07

    {$IFDEF VER150}
    ,Variants
    {$ENDIF}

    ;
{$ENDIF}

function YYLanceUtilisatRessMul (strArgs:string) : String;
begin
     Result := AGLLanceFiche ('YY','YYUTILISATRESSMUL', '', '', strArgs);
end;

procedure TOF_YYUTILISATRESSMUL.OnArgument (strArgs:string);
var
    tmp      : String;
begin
 inherited;
 bModeSelect := False;

 //--- Mot clé "SELECTION" pour utiliser la fiche en mode sélection
 tmp := ReadTokenSt(strArgs);
 While tmp<>'' do
  begin
   if tmp = 'SELECTION' then
        bModeSelect := True
   else if tmp = 'MULTISELECTION' then
   begin
        // $$$ JP 02/04/07: multi sélection
        bMultiSelect := TRUE;
      {$IFDEF EAGLCLIENT}
        THDBGrid(GetControl('FListe')).MultiSelect := True;
      {$ELSE}
        THDBGrid(GetControl('FListe')).MultiSelection := True;
      {$ENDIF}
        TToolBarButton97 (GetControl ('BOUVRIR')).OnClick := BMultiSelectOuvrirClick;
   end
   else if Copy (tmp, 1, 6) = 'USERS=' then
   begin
        // $$$ JP 02/04/07: restriction aux utilisateurs spécifiés
        System.Delete (tmp, 1, 6);
        SetControlText ('XX_WHERE', 'US_UTILISATEUR IN ("' + StringReplace (tmp, ',', '","', [rfIgnoreCase, rfReplaceAll]) + '")');
   end;

   tmp := ReadTokenSt (strArgs);
  end;

 //--- Création du menu PopUp
 PMUtilisatRess := TPopupMenu.Create (Ecran);
 PMUtilisatRess.Images  := FMenuDisp.SmallImages;
 PMUtilisatRess.OnPopup := OnPopupUtilisatRess;

 //--- Création des Items du menu PopUp
 if (VH_DP.SeriaMessagerie) then
  begin
   MnNouveauMessage:=TMenuItem.Create (Ecran);
   if (MnNouveauMessage<>nil) then
    begin
     MnNouveauMessage.Caption := '&Nouveau message';
     MnNouveauMessage.OnClick := mnNouveauMessageClick;
     mnNouveauMessage.ImageIndex := cImgMenuDispMail; // $$$ JP 23/05/07
     PMUtilisatRess.Items.Add (MnNouveauMessage);
    end;
  end;

 {$IFDEF BUREAU}
 // $$$ JP 23/05/07: ré-écriture, ajout 3 tel de l'utilisateur
 if VH_DP.ctiAlerte <> nil then
  begin
       // Téléphone ressource
       TelMenu := TMenuItem.Create (Ecran);
       with TelMenu do
       begin
            Name       := 'MTEL';
            Caption    := 'Appeler';
            ImageIndex := cImgMenuDispTel;
            OnClick    := AppelerClick;
            Visible    := FALSE;
       end;
       PMUTILISATRESS.Items.Add (TelMenu);

       // Telephone 1 utilisateur
       Tel1Menu := TMenuItem.Create (Ecran);
       with Tel1Menu do
       begin
            Name       := 'MTEL1';
            Caption    := 'Appeler';
            ImageIndex := cImgMenuDispTel;
            OnClick    := AppelerClick;
            Visible    := FALSE;
       end;
       PMUTILISATRESS.Items.Add (Tel1Menu);

       // Telephone 2 utilisateur
       Tel2Menu := TMenuItem.Create (Ecran);
       with Tel2Menu do
       begin
            Name       := 'MTEL2';
            Caption    := 'Appeler';
            ImageIndex := cImgMenuDispTel;
            OnClick    := AppelerClick;
            Visible    := FALSE;
       end;
       PMUTILISATRESS.Items.Add (Tel2Menu);

       // Telephone 3 utilisateur
       Tel3Menu := TMenuItem.Create (Ecran);
       with Tel3Menu do
       begin
            Name       := 'MTEL3';
            Caption    := 'Appeler';
            ImageIndex := cImgMenuDispTel;
            OnClick    := AppelerClick;
            Visible    := FALSE;
       end;
       PMUTILISATRESS.Items.Add (Tel3Menu);
  end;
 {$ENDIF}

 //--- Gestion des evenements
 Ecran.OnKeyDown:=Form_OnKeyDown;
 THDBGrid (GetControl ('FLISTE')).PopupMenu := PMUtilisatRess;
 if bModeSelect then THDBGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;

 // Noms des zones libres assistants
 AfficheLibTablesLibres (Self);
end;

procedure TOF_YYUTILISATRESSMUL.OnClose;
begin
//{$IFDEF BUREAU}
// FreeAndNil (Tel1Menu);
//{$ENDIF}

     FreeAndNil (PMUtilisatRess);
end;

// $$$ JP 23/05/07
{$IFDEF BUREAU}
procedure TOF_YYUTILISATRESSMUL.OnPopupUtilisatRess (Sender:TObject);
var
   vTel   :variant;
   strTel :string;
begin
     // Telephone Ressource
     if VH_DP.ctiAlerte <> nil then
     begin
      vTel := GetField ('ARS_TELEPHONE');
      if (VarIsEmpty (vTel) = FALSE) and (VarIsNull (vTel) = FALSE) then
          strTel := Trim (VarAsType (vTel, varString))
      else
          strTel := '';
      with TelMenu do
      begin
           if strTel <> '' then
           begin
                Caption := 'Appeler le ' + strTel;
                Visible := TRUE;
           end
           else
                Visible := FALSE;
      end;

      // Telephone 1 utilisateur
      vTel := GetField ('US_TEL1');
      if (VarIsEmpty (vTel) = FALSE) and (VarIsNull (vTel) = FALSE) then
          strTel := Trim (VarAsType (vTel, varString))
      else
          strTel := '';
      with Tel1Menu do
      begin
           if strTel <> '' then
           begin
                Caption := 'Appeler le ' + strTel;
                Visible := TRUE;
           end
           else
                Visible := FALSE;
      end;

      // Telephone 2 utilisateur
      vTel := GetField ('US_TEL2');
      if (VarIsEmpty (vTel) = FALSE) and (VarIsNull (vTel) = FALSE) then
          strTel := Trim (VarAsType (vTel, varString))
      else
          strTel := '';
      with Tel2Menu do
      begin
           if strTel <> '' then
           begin
                Caption := 'Appeler le ' + strTel;
                Visible := TRUE;
           end
           else
                Visible := FALSE;
      end;

      // Telephone 3 utilisateur
      vTel := GetField ('US_TEL3');
      if (VarIsEmpty (vTel) = FALSE) and (VarIsNull (vTel) = FALSE) then
          strTel := Trim (VarAsType (vTel, varString))
      else
          strTel := '';
      with Tel3Menu do
      begin
           if strTel <> '' then
           begin
                Caption := 'Appeler le ' + strTel;
                Visible := TRUE;
           end
           else
                Visible := FALSE;
      end;
     end;
end;

procedure TOF_YYUTILISATRESSMUL.AppelerClick (Sender:TObject);
var
   strTel :string;
begin
     if VH_DP.ctiAlerte <> nil then
     begin
          strTel := '';
          with Sender as TMenuItem do
          begin
               if Name = 'MTEL' then
                    strTel := Trim (GetField ('ARS_TELEPHONE'))
               else if Name = 'MTEL1' then
                    strTel := Trim (GetField ('US_TEL1'))
               else if Name = 'MTEL2' then
                    strTel := Trim (GetField ('US_TEL2'))
               else if Name = 'MTEL3' then
                    strTel := Trim (GetField ('US_TEL3'));
          end;

          if (strTel <> '') and (PgiAsk ('Appeler le ' + strTel + ' ?') = mrYes) then
             VH_DP.ctiAlerte.MakeCall (strTel);
     end;
end;

procedure TOF_YYUTILISATRESSMUL.FListe_OnDblClick(Sender: TObject);
begin
  if Not VarIsNull(GetField('US_UTILISATEUR')) then
    TFMul(Ecran).Retour := GetField('US_UTILISATEUR');
  Ecran.Close;
end;

//----------------------------------
//--- Nom : MnNouveauMessageClick
//----------------------------------
procedure TOF_YYUTILISATRESSMUL.mnNouveauMessageClick(Sender: TObject);
{ Var SSql, SEmail : String;
    RSql : TQuery; }
begin
 if (VarIsNull(GetField('US_UTILISATEUR'))) then Exit;

 //--- Récupération de l'adresse Email de l'utilisateur sélectionné
{ SSql:='SELECT US_EMAIL FROM UTILISAT WHERE US_UTILISATEUR="'+GetField ('US_UTILISATEUR')+'"';
 RSql:=OpenSql (SSql,True);
 if not (RSql.eof) then
  SEmail:=RSql.FindField ('US_EMAIL').AsString;
 Ferme (RSql); }

 //--- Envoit de l'Email
 if (VH_DP.SeriaMessagerie) then
  ShowNewMessage ('','','',GetField ('US_UTILISATEUR'))
end;

//---------------------------
//--- Nom : Form_OnKeyDown
//---------------------------
procedure TOF_YYUTILISATRESSMUL.Form_OnKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
begin
     case Key of
          VK_F11:
          begin
               PMUTILISATRESS.Popup (Mouse.CursorPos.X, Mouse.CursorPos.Y); // $$$ JP 16/05/06 FQ 11051: TPopupMenu(GetControl('PM_UTILISATRESS')).Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
               Key := 0;
          end;
     else
         TFMul (Ecran).FormKeyDown (Sender, Key, Shift);
     end;
end;
{$ENDIF} // BUREAU

procedure TOF_YYUTILISATRESSMUL.BMultiSelectOuvrirClick (Sender:TObject);
var
    Lst       :THDBGrid;
    iNbSel, i :integer;
    Q         :THQuery;
    strUsers  :string;
begin
     strUsers := '';
     Lst := TFMul(Ecran).FListe;
     Q   := TFMul(Ecran).Q;
     if (Lst.AllSelected = TRUE) or (Lst.NbSelected > 0) then
     begin
          // Traitement
          if Lst.AllSelected then
          begin
{$IFDEF EAGLCLIENT}
                if not TFMul(Ecran).FetchLesTous then
                   PGIInfo ('Impossible de récupérer tous les enregistrements')
                else
{$ENDIF}
                begin
                     Q.First;
                     while Not Q.EOF do
                     begin
                          if strUsers <> '' then
                             strUsers := strUsers + ',';
                          strUsers := strUsers + Q.FindField ('US_UTILISATEUR').AsString;

                          Q.Next;
                     end;
                end;
          end
          else
          begin
               iNbSel := Lst.NbSelected;
               for i := 0 to iNbSel-1 do
               begin
                    Lst.GotoLeBookmark (i);
{$IFDEF EAGLCLIENT}
                    Q.TQ.Seek(Lst.Row - 1) ;
{$ENDIF}
                    if strUsers <> '' then
                       strUsers := strUsers + ',';
                    strUsers := strUsers + Q.FindField ('US_UTILISATEUR').AsString;
               end;
          end;
     end;

     if strUsers <> '' then
     begin
          TFMul (Ecran).Retour := strUsers;
          Ecran.Close;
     end
     else
         PgiInfo ('Veuillez sélectionner au moins un utilisateur');
end;

initialization
  RegisterClasses ( [TOF_YYUTILISATRESSMUL] ) ;

end.

