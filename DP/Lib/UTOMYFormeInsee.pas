{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOM_JUFORMEJUR
Mots clefs ... : TOM;JUFORMEJUR
*****************************************************************}

Unit UTOMYFormeInsee;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
{$IFDEF EAGLCLIENT}
   UTob, eFiche,
{$ELSE}
   db,
{$ENDIF}
     UTOM, Classes, comctrls, SysUtils, hctrls, HTB97, controls,
     DpJurOutils, UDpJurOutilsFormeInsee, HMsgBox;

//////////////////////////////////////////////////////////////////
Type
   TOM_YFORMESINSEE  = Class (TOM)

         procedure OnArgument(sParams_p : String); override;
         procedure OnNewRecord; override ;
         procedure OnLoadRecord; override;
         procedure OnChangeField(F: TField); override ;
         procedure OnUpdateRecord; override ;
         procedure OnAfterUpdateRecord; override ;
         procedure OnDeleteRecord; override ;
         procedure OnClose; override;
      public

         procedure OnChange_Node(Sender : TObject; tnNode_p : TTreeNode);
         procedure OnClick_Expand(Sender : TObject);
         procedure OnClick_Collapse(Sender : TObject);
         procedure OnClick_CmbForme(Sender : TObject);
         procedure OnClick_CmbFormeDetail(Sender : TObject);
         procedure OnClick_CmbRegleFisc(Sender : TObject);

      private

         tvForme_c : TTreeView;
         tnNode_c  : TTreeNode;
         LaForme_c : TPForme;

         bClose_c : boolean;

         procedure TWPositionne(tnNode_p : TTreeNode);
         procedure FormeAfficheCache(sForme_p : string);
         procedure RegleFiscAfficheCache(sRegleFisc_p : string);

end ;

//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YFORMESINSEE.OnArgument (sParams_p : String);
begin
   Inherited ;
   bClose_c := false;

   SetFocusControl('FORMES_TV');

   tvForme_c := TTreeView(GetControl('FORMES_TV'));
   TWInitListeIcones(tvForme_c);
   TWChargeFormes(tvForme_c);
   TWPositionne(tvForme_c.Items.GetFirstNode);

   tvForme_c.OnChange := OnChange_Node;
   THDBValComboBox(GetControl('YFJ_FORME')).OnClick := OnClick_CmbForme;
   THDBValComboBox(GetControl('YFJ_FORMEGRPPRIVE')).OnClick := OnClick_CmbFormeDetail;
   THDBValComboBox(GetControl('YFJ_FORMESTE')).OnClick := OnClick_CmbFormeDetail;
   THDBValComboBox(GetControl('YFJ_FORMESCI')).OnClick := OnClick_CmbFormeDetail;
   THDBValComboBox(GetControl('YFJ_FORMEASSO')).OnClick := OnClick_CmbFormeDetail;

   THDBValComboBox(GetControl('YFJ_REGLEFISC')).OnClick := OnClick_CmbRegleFisc;
   TToolbarButton97(GetControl('BEXPAND')).OnClick := OnClick_Expand;
   TToolbarButton97(GetControl('BCOLLAPSE')).OnClick := OnClick_Collapse;

end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 05/10/02
Procédure .... : OnNewRecord
Description .. : Force le champ PREDEFINI avec 'STD' pour que les
                 données ne soient pas écrasées au rechargement de
                 socref
Paramètres ... :
*****************************************************************}

procedure TOM_YFORMESINSEE.OnNewRecord;
begin
   Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnLoadRecord;
begin
   Inherited ;
   SetControlEnabled('YFJ_NIVEAU', (DS.State = dsInsert));
   FormeAfficheCache(GetField('YFJ_FORME'));
   RegleFiscAfficheCache(GetField('YFJ_REGLEFISC'));
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnChangeField(F : TField);
begin
   Inherited ;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 05/10/02
Procédure .... : OnUpdateRecord
Description .. : En mode création, ajout d'un caractère spécial
                 au code (forcément STD)
Paramètres ... :
*****************************************************************}

procedure TOM_YFORMESINSEE.OnUpdateRecord ;
var
   iNiveau_l, iLgCle_l : integer;
   sCodeInsee_l : string;
begin
   Inherited ;
   if GetField('YFJ_CODEINSEE') = '' then
   begin
      LastError := 1 ;
      LastErrorMsg := 'La saisie d''un code INSEE est obligatoire...' ;
      SetFocusControl('YFJ_CODEINSEE');
      exit;
   end;

   if (GetField('YFJ_NIVEAU') = 0) then
   begin
      LastError := 1 ;
      LastErrorMsg := 'La saisie d''un niveau INSEE est obligatoire...' ;
      SetFocusControl('YFJ_NIVEAU');
      exit;
   end;

   iNiveau_l := GetField('YFJ_NIVEAU');
   sCodeInsee_l := GetField('YFJ_CODEINSEE');
   iLgCle_l := TWLongueurCodeInsee(iNiveau_l);
   if length(sCodeInsee_l) <> iLgCle_l then
   begin
      LastError := 1 ;
      LastErrorMsg := 'Le code INSEE de niveau ' + IntToStr(iNiveau_l) + ' doit être de ' + IntToStr(iLgCle_l) + ' caractères...' ;
      SetFocusControl('YFJ_NIVEAU');
      exit;
   end;

   if (GetField('YFJ_PREDEFINI') = '') then
   begin
      LastError := 1 ;
      LastErrorMsg := 'La saisie du prédéfini CEGID est obligatoire...' ;
      SetFocusControl('YFJ_PREDEFINI');
      exit;
   end;

   if (GetField('YFJ_NONACTIF') = '-') and (GetField('YFJ_FORME') = '') then
   begin
      LastError := 1 ;
      LastErrorMsg := 'La saisie d''une forme juridique est obligatoire pour un noeud actif...' ;
      SetFocusControl('YFJ_FORME');
      exit;
   end;

   if DS.State = dsInsert then exit;

   LaForme_c^.sCodeInsee_c  := GetField('YFJ_CODEINSEE');
   LaForme_c^.sForme_c      := GetField('YFJ_FORME');
   LaForme_c^.sFormePrive_c := GetField('YFJ_FORMEGRPPRIVE');
   LaForme_c^.sFormeSte_c   := GetField('YFJ_FORMESTE');
   LaForme_c^.sFormeSci_c   := GetField('YFJ_FORMESCI');
   LaForme_c^.sFormeAsso_c  := GetField('YFJ_FORMEASSO');
   LaForme_c^.sLibelle_c    := TWFormateLib(LaForme_c^.sForme_c,
                                            LaForme_c^.sCodeInsee_c,
                                            GetField('YFJ_LIBELLE'));
   LaForme_c^.iNiveau_c     := GetField('YFJ_NIVEAU');
   LaForme_c^.sNonActif_c   := GetField('YFJ_NONACTIF');
   LaForme_c^.sCoop_c       := GetField('YFJ_COOP');
   LaForme_c^.sRegleFisc_c  := GetField('YFJ_REGLEFISC');
   LaForme_c^.sSectionBnc_c := GetField('YFJ_SECTIONBNC');

end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnAfterUpdateRecord ;
begin
   Inherited ;
   if GetControl('FORMES_TV') <> nil then
   begin
      if DS.State = dsInsert Then
      begin
         TWChargeFormes(tvForme_c);
      end
      else
      begin
         tvForme_c.Items.BeginUpdate;
         tnNode_c.Data := LaForme_c;
         tnNode_c.Text := LaForme_c^.sLibelle_c;
         tnNode_c.ImageIndex := TWIcone(LaForme_c^.sNonActif_c, LaForme_c^.iNiveau_c);
         tnNode_c.SelectedIndex := TWIcone(LaForme_c^.sNonActif_c, LaForme_c^.iNiveau_c);
         tvForme_c.Items.EndUpdate;
      end;
   end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnClose;
begin
   Inherited;
   bClose_c := true;
   TWVideFormes(tvForme_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnDeleteRecord ;
begin
   Inherited ;
   TWChargeFormes(tvForme_c);
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOM_YFORMESINSEE.OnChange_Node(Sender : TObject; tnNode_p : TTreeNode);
begin
   if bClose_c then exit;
   if tnNode_p = nil then exit;

   TWPositionne(tnNode_p);
   {$IFDEF EAGLCLIENT}
   if DS.State in [dsInsert, dsEdit] then
   begin
      MessageAlerte(TFFiche(Ecran).HM.Mess[5]);

//      TFFiche(Ecran).Bouge(TNavigateBtn(nil));
//      if not TFFiche(Ecran).OnSauve then exit;
   end;
//   tvForme_c.Items
//   TFFiche(Ecran).QFiche.Seek()
   TFFiche(Ecran).FLequel := LaForme_c^.sCodeInsee_c;
   TFFiche(Ecran).ReloadDb;
   {$ELSE}
   DS.Locate('YFJ_CODEINSEE', LaForme_c^.sCodeInsee_c, []);
   {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YFORMESINSEE.OnClick_CmbForme(Sender : TObject);
begin
   if GetControlText('YFJ_FORME') = '' then
   begin
      ModeEdition(DS);
      SetField('YFJ_FORMEGRPPRIVE', '');
      SetField('YFJ_FORMESTE', '');
      SetField('YFJ_FORMESCI', '');
      SetField('YFJ_FORMEASSO', '');
   end;

   FormeAfficheCache(GetControlText('YFJ_FORME'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnClick_CmbFormeDetail(Sender : TObject);
var
   sControle_l : string;
begin
   sControle_l := TControl(Sender).Name;
   if GetControlText(sControle_l) <> '' then
   begin
      ModeEdition(DS);
      if sControle_l <> 'YFJ_FORMEGRPPRIVE' then
         SetField('YFJ_FORMEGRPPRIVE', '');
      if sControle_l <> 'YFJ_FORMESTE' then
         SetField('YFJ_FORMESTE', '');
      if sControle_l <> 'YFJ_FORMESCI' then
         SetField('YFJ_FORMESCI', '');
      if sControle_l <> 'YFJ_FORMEASSO' then
         SetField('YFJ_FORMEASSO', '');
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YFORMESINSEE.OnClick_CmbRegleFisc(Sender : TObject);
begin
   if GetControlText('YFJ_REGLEFISC') <> 'BNC' then
   begin
      ModeEdition(DS);
      SetField('YFJ_SECTIONBNC', '');
   end;

   RegleFiscAfficheCache(GetControlText('YFJ_REGLEFISC'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnClick_Expand(Sender : TObject);
begin
   OnExpand(tvForme_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.OnClick_Collapse(Sender : TObject);
begin
   OnCollapse(tvForme_c);
   OnChange_Node(TObject(tvForme_c), tvForme_c.Items.GetFirstNode);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.TWPositionne(tnNode_p : TTreeNode);
begin
   tnNode_c := tnNode_p;
   tnNode_c.Selected := true;
   LaForme_c := tnNode_c.Data;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.FormeAfficheCache(sForme_p : string);
begin
   SetControlEnabled('YFJ_FORMEGRPPRIVE', (sForme_p <> ''));
   SetControlEnabled('YFJ_FORMESTE', (sForme_p <> ''));
   SetControlEnabled('YFJ_FORMESCI', (sForme_p <> ''));
   SetControlEnabled('YFJ_FORMEASSO', (sForme_p <> ''));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YFORMESINSEE.RegleFiscAfficheCache(sRegleFisc_p : string);
begin
   SetControlEnabled('YFJ_SECTIONBNC', (sRegleFisc_p = 'BNC'));
end;


//////////////////////////////////////////////////////////////////

Initialization
   registerclasses ( [ TOM_YFORMESINSEE ] ) ;
end.
