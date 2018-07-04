{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
unit UTOFAnnuLienTri;
////////////////////////////////////////////////////////////////////////////////
interface
////////////////////////////////////////////////////////////////////////////////
uses
   {$IFNDEF EAGLCLIENT}
   FE_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
   {$ELSE}
   MaineAGL,
   {$ENDIF}
   Vierge, Controls, Classes, hmsgbox, SysUtils, Menus, StdCtrls,
   Hctrls, UCoherence, AnnOutils, DpJurOutils, CLASSAnnuLien, HTB97,
   UOutilsAnnuLienTri, UTOB, UTOF, UCLASSMyMenu, UCLASSMySubMenu, UCLASSKeys, UComOutils;
////////////////////////////////////////////////////////////////////////////////
type
   TOF_ANNULIEN_TRI = class(TOF)

      public
         procedure OnArgument(Arguments : String ) ; override ;
         procedure OnClose ; override ;

      protected
         procedure OnClick_ANNUAIRE(Sender: TObject);
         procedure OnClick_PopFonction(Sender: TObject);
         procedure OnClick_SupprimerLien(Sender: TObject);
         procedure OnClick_ConsulterLien(Sender: TObject);
         //
         procedure OnClick_ListeFonction(Sender: TObject);
         procedure OnRowEnter_ListeFonction(Sender: TObject; Ou: Integer;
                                             var Cancel: Boolean; Chg: Boolean);
         procedure OnDblClick_ListeLien(Sender: TObject);
         procedure OnChange_Typedos(Sender : TObject);

      private
         sGuidPerDos_c : string;
         sNoOrdre_c : string;
         sTypeDosCharge_c, sTypeDosAffiche_c : string;
         sForme_c : string;
         sCodeDos_c : string;
         sTypeFiche_c : string;
         {+GHA - 12/2007}
         sFicheName_c : string;
         {-GHA - 12/2007}

         OBFonctions_c, OBCurListeLiens_c, OBCurLien_c : TOB;
         sCodeFonction_c : string;
         sGuidPer_c : string;
         sRetour_c : string;

         MenuFonction_c : TMySubMenu;
         MyKeys_c : TKeysPlusVie;
         //
         procedure AfficheEnteteLienDos;
         function OuvreFicheLien(sGuidPerDos_p, sTypeDos_p, sNoOrdre_p, sCodeFonction_p, sGuidPer_p, sForme_p, sAction_p : string) : string;
////////////////////////////////////////////////////////////////////////////////
end;
////////////////////////////////////////////////////////////////////////////////

implementation
////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnArgument(Arguments : String ) ;
var
 bFSynthese : boolean;
begin
   sGuidPerDos_c := ReadTokenSt(Arguments);
   sTypeDosCharge_c := ReadTokenSt(Arguments);
   sNoOrdre_c    := ReadTokenSt(Arguments);
   sForme_c      := ReadTokenSt(Arguments);
   sCodeDos_c    := ReadTokenSt(Arguments);
   sTypeFiche_c  := ReadTokenSt(Arguments);
   {+GHA - 12/2007}
   sFicheName_c  := ReadTokenSt(Arguments);
   {-GHA - 12/2007}

   sRetour_c := '';

   THGrid(GetControl('LISTEFONCTION')).OnRowEnter := OnRowEnter_ListeFonction;
   THGrid(GetControl('LISTEFONCTION')).OnClick := OnClick_ListeFonction;
   THGrid(GetControl('LISTELIEN')).OnDblClick := OnDblClick_ListeLien;
   TButton(GetControl('CONSULTERLIEN')).OnClick := OnClick_ConsulterLien;
   TButton(GetControl('SUPPRIMERLIEN')).OnClick := OnClick_SupprimerLien;
   TButton(GetControl('ACCEDERFICHE')).OnClick := OnClick_ANNUAIRE;

   MyKeys_c := TKeysPlusVie.Create(TFVierge(Ecran),
               [vk_Ins_c, vk_Sup_c, vk_Ouv_c],
               [TToolbarButton97(GetControl('INSERERLIEN')),
                TToolbarButton97(GetControl('SUPPRIMERLIEN')),
                TToolbarButton97(GetControl('CONSULTERLIEN'))]);
   Ecran.OnKeyDown := MyKeys_c.OnVieKeyDown;

  {// FQ 11763 (KPMG) : mise en commentaire car on détaille tout le temps les droits
   // Pas le droit de modifier
   if Not JaiLeDroitConceptBureau(ccModifAnnuaireOuLiens) then
      begin
      // pas de création/suppression
      SetControlVisible('INSERERLIEN', False);
      SetControlVisible('SUPPRIMERLIEN', False);
      end

   // ou bien droits plus détaillés
   else }

  {+GHA - 12/2007 concept KPMG}
   if sFicheName_c = 'FSynthese' then
      bFSynthese := TRUE
   else
      bFSynthese := FALSE;

   if ((bFSynthese) and (not JaiLeDroitConceptBureau(ccModifAnnuaireDossier))) or
      (not bFSynthese) then
   begin
      {-GHA - 12/2007 concept KPMG}
      // Pas de création
      if Not JaiLeDroitConceptBureau(ccCreatAnnuaireOuLiens) then
      begin
         SetControlVisible('INSERERLIEN', False);
         // Désactivation du clic-droit
         THGrid(GetControl('LISTEFONCTION')).PopupMenu := Nil;
      end;

      // Pas de suppression
      if Not JaiLeDroitConceptBureau(ccSupprAnnuaireOuLiens) then
         SetControlVisible('SUPPRIMERLIEN', False);
   end;

   THValComboBox(GetControl('JFT_CATFONCT')).OnClick := OnChange_Typedos;
   if sTypeDosCharge_c <> '' then
   begin
      THValComboBox(GetControl('JFT_CATFONCT')).Plus := 'AND CO_CODE = "' + sTypeDosCharge_c + '"';
      THValComboBox(GetControl('JFT_CATFONCT')).Value := sTypeDosCharge_c;
      SetControlEnabled('JFT_CATFONCT', false);
   end
   else if sTypeFiche_c = 'JUR' then
      THValComboBox(GetControl('JFT_CATFONCT')).Value := 'STE';

   AfficheEnteteLienDos;
   ChargeFonctionsEtLiens(sGuidPerDos_c, sNoOrdre_c, sTypeDosCharge_c, sTypeFiche_c, sForme_c,
                          OBFonctions_c, OBCurListeLiens_c, OBCurLien_c);

   MenuFonction_c := TMySubMenu.Create(TPopupMenu(GetControl('POPFONCTIONS')),
                                       THValComboBox(GetControl('JFT_CATFONCT')));
   MenuFonction_c.SubMenuAdd(OBFonctions_c, 'JTF_FONCTABREGE', 'JTF_FONCTION', 'JFT_TYPEDOS', OnClick_PopFonction);

   OnChange_Typedos(nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 06/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnChange_Typedos(Sender : TObject);
begin
   sTypeDosAffiche_c := THValComboBox(GetControl('JFT_CATFONCT')).Value;
      
   AfficheFonctionsEtLiens(sTypeDosAffiche_c,
                           THGrid(GetControl('LISTEFONCTION')), THGrid(GetControl('LISTELIEN')),
                           OBFonctions_c, OBCurListeLiens_c, OBCurLien_c);
   TFVierge(Ecran).HMTrad.ResizeGridColumns(THGrid(GetControl('LISTEFONCTION')));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnClose;
begin
   inherited;
   Verifcoherence(sGuidPerDos_c, sCodeDos_c);
   TFVierge(Ecran).FRetour := sRetour_c;
   //
   OBFonctions_c.ClearDetail;
   OBFonctions_c.Free;
   if MyKeys_c <> nil then
      MyKeys_c.Free;
   if MenuFonction_c <> nil then
      FreeAndNil(MenuFonction_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnClick_ANNUAIRE(Sender: TObject);
var
  sCode_l : string;
begin
  OBCurLien_c := SelectCurLien(THgrid(GetControl('LISTELIEN')));
  if OBCurLien_c = nil then exit;
  sCode_l := OBCurLien_c.GetValue('ANL_GUIDPER');
  AGLLanceFiche('YY','ANNUAIRE', sCode_l, sCode_l,';;;' + sCode_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnClick_ConsulterLien(Sender: TObject);
begin
  OnDblClick_ListeLien(Sender);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnClick_ListeFonction(Sender: TObject);
begin
   AfficheCurLiens(THGrid(GetControl('LISTEFONCTION')), THGrid(GetControl('LISTELIEN')),
                           OBCurListeLiens_c, OBCurLien_c);
   TFVierge(Ecran).HMTrad.ResizeGridColumns(THGrid(GetControl('LISTELIEN')));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ANNULIEN_TRI.AfficheEnteteLienDos;
//var
//  OBJur_l : TOB;
begin
{   SetControlVisible('PDOSSIER', (sTypeFiche_c <> 'DP'));
   if sTypeFiche_c <> 'DP' then
   begin
      OBJur_l := TOB.Create('JURIDIQUE', nil, -1);
      OBJur_l.SelectDB('"' + sGuidPerDos_c + '";"' + sTypeDos_c + '";' + sNoOrdre_c, nil);
      // $$$ JP 04/09/06 - plus autorisé en D7/Unicode ... THEdit(GetControl('JUR_NOMDOS')).Text := OBJur_l.GetValue('JUR_NOMDOS');
                                                       //  THEdit(GetControl('JUR_DOSLIBELLE')).Text := OBJur_l.GetValue('JUR_DOSLIBELLE');
      SetControlText ('JUR_NOMDOS',     OBJur_l.GetValue('JUR_NOMDOS'));
      SetControlText ('JUR_DOSLIBELLE', OBJur_l.GetValue('JUR_DOSLIBELLE'));

      OBJur_l.Free;
   end;}
   THValComboBox(GetControl('JUR_FORME')).Value := sForme_c;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnRowEnter_ListeFonction(Sender: TObject; Ou: Integer;
                                                    var Cancel: Boolean; Chg: Boolean);
begin
   AfficheCurLiens(THGrid(GetControl('LISTEFONCTION')), THGrid(GetControl('LISTELIEN')),
                   OBCurListeLiens_c, OBCurLien_c);
   TFVierge(Ecran).HMTrad.ResizeGridColumns(THGrid(GetControl('LISTELIEN')));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnClick_SupprimerLien(Sender: TObject);
var
  Texte1, sTypeDos_l : string;
  iFRow_l, iLRow_l : integer;
begin
   OBCurLien_c := SelectCurLien(THgrid(GetControl('LISTELIEN')));
   if OBCurLien_c = nil then
      exit;

   iFRow_l := THgrid(GetControl('LISTEFONCTION')).Row;
   iLRow_l := THGrid(GetControl('LISTELIEN')).Row;

   sCodeFonction_c := OBCurLien_c.GetValue('FONCTION');
   sGuidPer_c := OBCurLien_c.GetValue ('ANL_GUIDPER') ;
   sTypeDos_l := OBCurLien_c.GetValue ('ANL_TYPEDOS') ;

   Texte1 := 'Attention : Vous allez supprimer ce lien.#13#10Confirmez-vous l''opération ?';
   if PgiAsk( Texte1, 'Suppression des liens' ) = MrYes then
   begin
      SupprimeUnLien(1, sGuidPerDos_c, sGuidPer_c, sTypeDos_l, sCodeFonction_c);
      OBCurLien_c.Free;
      OBCurLien_c := nil;
      sRetour_c := 'X';
   end;

   ChargeFonctionsEtLiens(sGuidPerDos_c, sNoOrdre_c, sTypeDosCharge_c, sTypeFiche_c, sForme_c,
                           OBFonctions_c, OBCurListeLiens_c, OBCurLien_c);
   AfficheFonctionsEtLiens(sTypeDosAffiche_c,
                           THGrid(GetControl('LISTEFONCTION')), THGrid(GetControl('LISTELIEN')),
                           OBFonctions_c, OBCurListeLiens_c, OBCurLien_c);

   THgrid(GetControl('LISTEFONCTION')).Row := iFRow_l;
   OnClick_ListeFonction(TObject(GetControl('LISTEFONCTION')));

   if THGrid(GetControl('LISTELIEN')).RowCount <= iLRow_l then
      iLRow_l := THGrid(GetControl('LISTELIEN')).RowCount - 1;
   THGrid(GetControl('LISTELIEN')).Row := iLRow_l;

end;




{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnDblClick_ListeLien(Sender: TObject);
var
   iLigFnc_l, iLigLns_l : integer;
   sRes_l, sTypeDos_l : string;
begin
   OBCurLien_c := SelectCurLien(THgrid(GetControl('LISTELIEN')));
   if OBCurLien_c = nil then exit;
   
   sCodeFonction_c := OBCurLien_c.GetValue('FONCTION');
   sGuidPer_c := OBCurLien_c.GetValue ('ANL_GUIDPER') ;
   sTypeDos_l := OBCurLien_c.GetValue ('ANL_TYPEDOS') ;

   sRes_l := OuvreFicheLien(sGuidPerDos_c, sTypeDos_l, sNoOrdre_c, sCodeFonction_c, sGuidPer_c, sForme_c, 'ACTION=MODIFICATION');

   if (sRes_l <> '') then   // on recharge les liens
   begin
      iLigFnc_l := THgrid(GetControl('LISTEFONCTION')).Row;
      iLigLns_l := THgrid(GetControl('LISTELIEN')).Row;

      ReChargeCurLiens(sGuidPerDos_c, sNoOrdre_c, '',
                       sTypeFiche_c, sForme_c, sCodeFonction_c,
                       OBFonctions_c);

      THgrid(GetControl('LISTEFONCTION')).Row := iLigFnc_l;
      OnClick_ListeFonction(TObject(GetControl('LISTEFONCTION')));
      if THgrid(GetControl('LISTELIEN')).RowCount <= iLigLns_l then
         THgrid(GetControl('LISTELIEN')).Row := THgrid(GetControl('LISTELIEN')).RowCount - 1;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ANNULIEN_TRI.OnClick_PopFonction(Sender : TObject);
var
   OldFonction, sGuidPerSel_l, sTypeDos_l : string;
   iFRow_l : integer;
begin
   OBCurListeLiens_c := SelectCurListeLiens(THgrid(GetControl('LISTEFONCTION')));
   if OBCurListeLiens_c <> nil then
      OldFonction := OBCurListeLiens_c.GetValue('JTF_FONCTION')
   else
      OldFonction := sCodeFonction_c;

   sCodeFonction_c := MenuFonction_c.GetHint;
   sTypeDos_l := MenuFonction_c.GetSubLeHint;
   iFRow_l := THgrid(GetControl('LISTEFONCTION')).Row;

   sGuidPerSel_l := AGLLanceFiche('JUR','ANNUAIRE_SEL','','','');
   if (sGuidPerSel_l <> '0') and (sGuidPerSel_l <> '') then
   begin
      sGuidPerSel_l := OuvreFicheLien(sGuidPerDos_c, sTypeDos_l, sNoOrdre_c, sCodeFonction_c, sGuidPerSel_l, sForme_c, 'ACTION=CREATION');

      if (sGuidPerSel_l <> '0') and (sGuidPerSel_l <> '') then // Recharge les liens modifiés
      begin
         sRetour_c := 'X';
         if (sTypeDos_l = 'STE') and (sCodeFonction_c <> 'INT') and (sCodeFonction_c <> 'TRS') then
         begin
            CreationLienPerDosVirtuel(sGuidPerDos_c, sTypeDos_l, 'INT', sGuidPerSel_l,
                                      StrToInt(sNoOrdre_c), sForme_c, sCodeDos_c);
         end;

         ChargeFonctionsEtLiens(sGuidPerDos_c, sNoOrdre_c, sTypeDosCharge_c, sTypeFiche_c, sForme_c,
                           OBFonctions_c, OBCurListeLiens_c, OBCurLien_c);
         AfficheFonctionsEtLiens(sTypeDosAffiche_c,
                           THGrid(GetControl('LISTEFONCTION')), THGrid(GetControl('LISTELIEN')),
                           OBFonctions_c, OBCurListeLiens_c, OBCurLien_c);

         THgrid(GetControl('LISTEFONCTION')).Row := iFRow_l;
         OnClick_ListeFonction(TObject(GetControl('LISTEFONCTION')));

         OBCurLien_c := OBCurListeLiens_c.FindFirst(['ANL_GUIDPER'],[sGuidPerSel_l],true);
         if OBCurLien_c <> nil then
            THgrid(GetControl('LISTELIEN')).Row := OBCurLien_c.GetIndex + 1;
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/02/2008
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
function TOF_ANNULIEN_TRI.OuvreFicheLien(sGuidPerDos_p, sTypeDos_p, sNoOrdre_p, sCodeFonction_p, sGuidPer_p, sForme_p, sAction_p : string) : string;
var
   sCle_l, sRange_l, sGuidPerSel_l : string;
begin
   sCle_l := sGuidPerDos_p + ';' + sTypeDos_p + ';' + sNoOrdre_p + ';' + sCodeFonction_p + ';' + sGuidPer_p;// + ';' + sForme_c;

   if sAction_p = 'ACTION=MODIFICATION' then
      sRange_l := sCle_l;

   if (sTypeDos_p = 'STE') then
   begin
      if (sCodeFonction_p = 'INT') and (sAction_p = 'ACTION=MODIFICATION') then
         sGuidPerSel_l := AGLLanceFiche('JUR', 'INTERVENANT', '', '', sAction_p + ';' + sCle_l + ';' + sForme_c)
      else
         sGuidPerSel_l := AGLLanceFiche('YY', 'FICHELIEN', sRange_l, sCle_l, sAction_p + ';' + sCle_l + ';' + sForme_c);
   end
   else
   begin
      sGuidPerSel_l := AGLLanceFiche('DP', 'FICHINTERVENTION', sRange_l, sCle_l, sAction_p + ';' + sCle_l + ';' + sForme_c);
   end;
   result := sGuidPerSel_l;
end;
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization
   registerclasses([TOF_ANNULIEN_TRI]) ;
end.
