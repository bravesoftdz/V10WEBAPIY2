{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOM_JUFORMEJUR
Mots clefs ... : TOM;JUFORMEJUR
*****************************************************************}

Unit UTOFYFormeInsee;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
{$IFDEF EAGLCLIENT}
   UTob,
{$ELSE}
   db,
{$ENDIF}
     Classes, comctrls, SysUtils, HTB97, Dialogs, StrUtils,
     UTOF, Vierge, UDpJurOutilsFormeInsee ,HMsgBox;

//////////////////////////////////////////////////////////////////
Type
   TOF_YFORMESINSEE  = Class (TOF)

         procedure OnArgument(sParams_p : String); override;
         procedure OnLoad; override;
         procedure OnClose; override;
      public

         procedure OnChange_Node(Sender : TObject; tnNode_p : TTreeNode);
         procedure OnClick_Expand(Sender : TObject);
         procedure OnClick_Collapse(Sender : TObject);
         procedure OnClick_BValider(Sender : TObject);
         procedure OnClick_BRechercher(Sender : TObject); // Gha 01/2008
         procedure FindDiag_OnFind(Sender : TObject);
         procedure OnKeyDown_Ecran(Sender: TObject; var Key: Word; Shift: TShiftState);
      private

         tvForme_c : TTreeView;
         tnNode_c : TTreeNode;
         LaForme_c : TPForme;
         bClose_c : boolean;
         FindDiag : TFindDialog;
         BtnRech : TToolbarButton97;

         procedure TWSelectionne(tnNode_p : TTreeNode);
         procedure TWRecherche(sCodeInsee_p : string);
         function FindStringMode(Node : TTreeNode ;FindTxt : string ; WholeWord,MatchCase : boolean):boolean;
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
procedure TOF_YFORMESINSEE.OnArgument(sParams_p : String);
begin
   Inherited ;
   bClose_c := false;
   TFVierge(Ecran).Retour := 'FALSE';

   SetControlEnabled('BVALIDER', false);
   TToolbarButton97(GetControl('BVALIDER')).OnClick := OnClick_BValider;
   TToolbarButton97(GetControl('BEXPAND')).OnClick := OnClick_Expand;
   TToolbarButton97(GetControl('BCOLLAPSE')).OnClick := OnClick_Collapse;

   {+GHA 01/2008 - FQ 11914 fonction de recherche dans l'arborescence}
   BtnRech := TToolbarButton97(GetControl('BRECHERCHER'));
   if BtnRech <> nil then
   begin
     BtnRech.OnClick := OnClick_BRechercher;
     // Raccourci clavier pour la fonction de recherche
     Ecran.OnKeyDown := OnKeyDown_Ecran;
   end;
   {-GHA 01/2008 - FQ 11914}

   SetFocusControl('FORMES_TV');
   tvForme_c := TTreeView(GetControl('FORMES_TV'));
   TWInitListeIcones(tvForme_c);
   TWChargeFormes(tvForme_c);
   if sParams_p <> '' then
   begin
      TWRecherche(sParams_p);
   end
   else
   begin
      TWSelectionne(tvForme_c.Items.GetFirstNode);
   end;
   tvForme_c.OnChange := OnChange_Node;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_YFORMESINSEE.OnLoad;
begin
   Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YFORMESINSEE.OnClose;
begin
   Inherited ;
   bClose_c := true;
   TWVideFormes(tvForme_c);
   FreeAndNil(FindDiag);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_YFORMESINSEE.OnChange_Node(Sender : TObject; tnNode_p : TTreeNode);
begin
   if bClose_c then exit;

   if tnNode_p = nil then
   begin
      SetControlEnabled('BVALIDER', false);
      exit;
   end;

   TWSelectionne(tnNode_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 16/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YFORMESINSEE.OnClick_BValider(Sender : TObject);
begin
   TFVierge(Ecran).Retour := LaForme_c^.sCodeInsee_c + ';' +
                             LaForme_c^.sForme_c + ';' +
                             LaForme_c^.sFormePrive_c + ';' +
                             LaForme_c^.sFormeSte_c + ';' +
                             LaForme_c^.sFormeSci_c + ';' +
                             LaForme_c^.sFormeAsso_c + ';' +
                             LaForme_c^.sCoop_c + ';' +
                             LaForme_c^.sRegleFisc_c + ';' +
                             LaForme_c^.sSectionBnc_c;

   TFVierge(Ecran).BValiderClick(Sender);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YFORMESINSEE.TWSelectionne(tnNode_p : TTreeNode);
begin
   tnNode_c := tnNode_p;
   LaForme_c := tnNode_c.Data;
   SetControlEnabled('BVALIDER', (LaForme_c^.sNonActif_c = '-'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_YFORMESINSEE.TWRecherche(sCodeInsee_p : string);
var
   bOK_l : boolean;
   sCodeNiv_l : string;
begin
   bOK_l := false;
   tnNode_c := tvForme_c.Items.GetFirstNode;

   while (tnNode_c <> nil) and not bOK_l do
   begin
      tnNode_c.Selected := true;
      LaForme_c := tnNode_c.Data;
      bOK_l := (LaForme_c^.sCodeInsee_c = sCodeInsee_p);

      if not bOK_l then
      begin
         sCodeNiv_l := Copy(sCodeInsee_p, 1, TWLongueurCodeInsee(LaForme_c^.iNiveau_c));
         if (LaForme_c^.sCodeInsee_c <> sCodeNiv_l) then
            tnNode_c := tnNode_c.getNextSibling
         else
            tnNode_c := tnNode_c.getNext;
      end;
   end;

   if bOK_l then
      SetControlEnabled('BVALIDER', (LaForme_c^.sNonActif_c = '-'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_YFORMESINSEE.OnClick_Expand(Sender : TObject);
begin
   OnExpand(tvForme_c);
end;

procedure TOF_YFORMESINSEE.OnClick_Collapse(Sender : TObject);
begin
   OnCollapse(tvForme_c);
   OnChange_Node(TObject(tvForme_c), tvForme_c.Items.GetFirstNode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : GHA
Créé le ...... : 03/01/2008
Modifié le ... : 03/01/2008
Description .. : Fonction de recherche dans l'arborescence
Mots clefs ... :
*****************************************************************}
{+GHA 01/2008 - FQ 11914}
procedure TOF_YFORMESINSEE.OnClick_BRechercher(Sender: TObject);
begin
  FindDiag := TFindDialog.Create(ecran);
  if FindDiag <> nil then
  begin
    FindDiag.OnFind := FindDiag_OnFind;
    FindDiag.Execute;
  end;
end;
{------------------------------------------------------------------------------}
procedure TOF_YFORMESINSEE.FindDiag_OnFind(Sender: TObject);
var
  text : string;
  i,start : integer;
  node : TTreeNode;
  bDown,bMatchCase,bWholeWord,bFindEnd : boolean;
begin
  bDown      := FALSE;
  bMatchCase := FALSE;
  bWholeWord := FALSE;
  bFindEnd   := TRUE;

  // info de la boite de dialogue Rechercher.
  with (sender as TFindDialog) do
  begin
    text := FindText;

    if frDown in Options then
      bDown := TRUE;

    if frMatchCase in Options then
      bMatchCase := TRUE;

    if frWholeWord in Options then
      bWholeWord := TRUE;
  end;

  //récupérer le focus sur le ctrl TV.
  if not tvForme_c.Focused then
    tvForme_c.SetFocus;

  node := tvForme_c.Selected;

  if node = nil then
    exit;

  //récupére la position de l'item dans la vue.
  start := node.AbsoluteIndex;

  if start > 0 then
  begin
    if bDown then
      inc(start)
    else
      Dec(start);
  end;

  if bDown then
  begin
    // Recherche descendante dans l'arborescence
    for i:=start to tvForme_c.Items.Count-1 do
    begin
      node := tvForme_c.Items[i];

      if FindStringMode(node,text,bWholeWord,bMatchCase) then
      begin
        bFindEnd := FALSE;
        break;
      end;
    end;
  end
  else
  begin
    // Recherche ascendante dans l'arborescence
    for i:=start downto 0 do
    begin
      node := tvForme_c.Items[i];

      if FindStringMode(node,text,bWholeWord,bMatchCase) then
      begin
        bFindEnd := FALSE;
        break;
      end;
    end;
  end;

  if bFindEnd then
    PGIINFO('Recherche terminée',Ecran.Caption);
end;
{------------------------------------------------------------------------------}
function TOF_YFORMESINSEE.FindStringMode(Node: TTreeNode; FindTxt: string;
  WholeWord, MatchCase: boolean): boolean;
var
  text : string;
begin
  result := FALSE;

  if (WholeWord) and (not MatchCase) then // Mot entier
  begin
    //vérif si le libellé contient l'expression recherchée.
    if pos(LowerCase(FindTxt),LowerCase(node.Text)) > 0 then
    begin
      //retrait du code insee dans la chaine de carac.
      text := Trim(copy(Node.Text,pos(':',Node.Text)+1,length(Node.Text)-pos(':',Node.Text)+1));

      if SameText(FindTxt,LeftStr(Text,length(FindTxt))) then
      begin
        node.Selected := TRUE;
        result := TRUE;
      end;
    end;
  end
  else if (not WholeWord) and (MatchCase) then // Distinction maj/min
  begin
    if pos(FindTxt,node.Text) > 0 then
    begin
      node.Selected := TRUE;
      result := TRUE;
    end;
  end
  else if (WholeWord) and (MatchCase) then // Mot entier + Distinction maj/min
  begin
    //vérif si le libellé contient l'expression recherchée.
    if pos(LowerCase(FindTxt),LowerCase(node.Text)) > 0 then
    begin
      //retrait du code insee dans la chaine de carac.
      text := Trim(copy(Node.Text,pos(':',Node.Text)+1,length(Node.Text)-pos(':',Node.Text)+1));

      if CompareStr(FindTxt,LeftStr(Text,length(FindTxt))) = 0 then
      begin
        node.Selected := TRUE;
        result := TRUE;
      end;
    end;
  end
  else // aucune option cochée
  begin
    if pos(LowerCase(FindTxt),LowerCase(node.Text)) > 0 then
    begin
      node.Selected := TRUE;
      result := TRUE;
    end;
  end;
end;
{------------------------------------------------------------------------------}
procedure TOF_YFORMESINSEE.OnKeyDown_Ecran(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    // CTRL + F
    70: if Shift = [ssCtrl] then
          BtnRech.Click;
  end;
end;
{-GHA 01/2008 - FQ 11914}

//////////////////////////////////////////////////////////////////
Initialization
   registerclasses ( [ TOF_YFORMESINSEE ] ) ;
end.
