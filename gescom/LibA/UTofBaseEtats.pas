unit UTofBaseEtats;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,AGLInit,SaisUtil,
{$IFDEF EAGLCLIENT}
   eMul,eQrs1,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,Mul,Qrs1,
{$ENDIF}
      HPanel,uTofAfBaseCodeAffaire,
      HCtrls,HEnt1,HMsgBox,UTOF, FactUtil,
      AfUtilArticle,
      TraducAffaire,utob,
      grids, Dicobtp, HSysMenu,HTB97,EntGC,confidentaffaire,
      uafo_ressource,ParamSoc;

(******************************************************************************)
// TOF_BASE_ETATS : objet TOF de gestion d'écran de lancement d'états
//
// Utilisation :  - Soit en tant que TOF associée à un écran de lancement d'état
//                - Soit en tant qu'ancêtre d'une TOF associée à un écran de lancement d'état
//
// Effet : - prise en compte de la gestion d'un intervalle de saisie d'affaires avec ou sans Tiers
//          (gestion des boutons, des combos, de la liste). Si le nom des champs a le préfixe 'ACT_'
//          la prise en compte est automatique, sinon, il faut dériver la TOF et créer la dérivée de la fonction
//          NomsChampsAffaire() qui renvoit les THEdit de gestion de l'affaire
//         - prise en compte de la gestion de 5 combos de saisie de rupture et des 5 check box de total associés s'ils
//          existent. (gestion du grisage, degrisage des champs suivant saisie)
//
// Auteur : Pierre Lenormand le 26/04/00
(******************************************************************************)
Type
     TOF_BASE_ETATS = Class (TOF_AFBASECODEAFFAIRE)
      procedure OnArgument(stArgument : String ) ; override ;
      procedure OnLoad ; override ;
      procedure GereRuptureComboCheck(ComboA:THValComboBox; CheckA,SautA:TCheckBox; ComboB:THValComboBox);
      procedure ComboRupture1Change(Sender: TObject); virtual ;
      procedure ComboRupture2Change(Sender: TObject); virtual ;
      procedure ComboRupture3Change(Sender: TObject); virtual ;
      procedure ComboRupture4Change(Sender: TObject); virtual ;
      procedure ComboRupture5Change(Sender: TObject); virtual ;
      procedure NomsChampsAffaireString(var sAff0, sAff1, sAff2, sAff3, sAff4 : string); virtual;
      procedure NomsChampsRessource(var Res1, Res2:THEdit; var LblRes1, LblRes2:THLabel);
      procedure TableauObjectsInvisibles(Crit: string; var iNbChamps:integer; var tbChamps:PString); virtual;
      function ListeObjectsInvisibles(Crit:string):TStringList;
     public
      UniteConv      :  THValComboBox;
      ComboRupture1  :  THValComboBox;
      ComboRupture2  :  THValComboBox;
      ComboRupture3  :  THValComboBox;
      ComboRupture4  :  THValComboBox;
      ComboRupture5  :  THValComboBox;
      ComboOrderBy   :  THValComboBox;
      sOrderByDefaut :  string;
      Saut1  :  TCheckBox;   //mcd 16/10/02 ajout zone saut associée
      Saut2  :  TCheckBox;   //mcd 16/10/02 ajout zone saut associée
      Saut3  :  TCheckBox;   //mcd 16/10/02 ajout zone saut associée
      Saut4  :  TCheckBox;   //mcd 16/10/02 ajout zone saut associée
      Saut5  :  TCheckBox;   //mcd 16/10/02 ajout zone saut associée
      CheckTotRupt1  :  TCheckBox;
      CheckTotRupt2  :  TCheckBox;
      CheckTotRupt3  :  TCheckBox;
      CheckTotRupt4  :  TCheckBox;
      CheckTotRupt5  :  TCheckBox;
      IsGestionAff   :  TCheckBox;
      RuptAff1       :  TCheckBox;
      RuptAff2       :  TCheckBox;
      RuptAff3       :  TCheckBox;                            
      RuptAff4       :  TCheckBox;
      RuptAff5       :  TCheckBox;
      Ressource1     :  THEdit;
      Ressource2     :  THEdit;
      LblRessource1  :  THLabel;
      LblRessource2  :  THLabel;

      procedure DeleteItemDansCombo(Combo:THValComboBox; Item:string);
      procedure CocherSiValeurDansCombo(ACocher:TCheckBox; sValeur:string; Combo:THValComboBox);
     END ;


implementation

procedure TOF_BASE_ETATS.OnArgument(stArgument : String );
var
  sAff0, sAff1, sAff2, sAff3, sAff4 : string;
  LOI       : TStringList;
  i         : integer;
  ThvalMul  : THMultiValComboBox;

  //cpt : string;
  //AFORessources : TAFO_Ressources;
  //ii,jj,posit:integer  ;
  //Thval:THVALComboBox;
  //text:string;

begin
  // héritage de la gestion du code affaire
  Inherited;
  // Drapeaux de présence de combos de ruptures à certaines valeurs
  RuptAff1 := TCheckBox(GetControl('RUPTAFF1'));
  RuptAff2 := TCheckBox(GetControl('RUPTAFF2'));
  RuptAff3 := TCheckBox(GetControl('RUPTAFF3'));
  RuptAff4 := TCheckBox(GetControl('RUPTAFF4'));
  RuptAff5 := TCheckBox(GetControl('RUPTAFF5'));

  { fait dans traducchamplibre qui est toujours ancêtre !!! 19/10/01
  // Gestion du champ indicatif du contexte gestion interne ou gestion d'affaire
  IsGestionAff := TCheckBox(GetControl('ISGESTAFF'));
  if (IsGestionAff<>nil) then
   ?? ne marche pas .... // mcd 26/09/01  if (ctxScot in V_PGI.PGIContexte)  then IsGestionAff.Checked:=false else IsGestionAff.Checked:=true;
     if (GetParamSoc('SO_AfFormatExer')<>'AUC') then IsGestionAff.Checked:=false else IsGestionAff.Checked:=true;
  }

  // Gestion du champ unité de conversion
  UniteConv := THValComboBox(GetControl('XX_VARIABLE3'));
  if (UniteConv<>nil) then UniteConv.Value:= VH_GC.AFMESUREACTIVITE;

  // Gestion des champs de selection de rupture
  ComboRupture1 := THValComboBox(GetControl('XX_RUPTURE1'));
  ComboRupture2 := THValComboBox(GetControl('XX_RUPTURE2'));
  ComboRupture3 := THValComboBox(GetControl('XX_RUPTURE3'));
  ComboRupture4 := THValComboBox(GetControl('XX_RUPTURE4'));
  ComboRupture5 := THValComboBox(GetControl('XX_RUPTURE5'));
  
  CheckTotRupt1 := TCheckBox(GetControl('TOTAL1'));
  CheckTotRupt2 := TCheckBox(GetControl('TOTAL2'));
  CheckTotRupt3 := TCheckBox(GetControl('TOTAL3'));
  CheckTotRupt4 := TCheckBox(GetControl('TOTAL4'));
  CheckTotRupt5 := TCheckBox(GetControl('TOTAL5'));

  Saut1 := TCheckBox(GetControl('SAUT1'));
  Saut2 := TCheckBox(GetControl('SAUT2'));
  Saut3 := TCheckBox(GetControl('SAUT3'));
  Saut4 := TCheckBox(GetControl('SAUT4'));
  Saut5 := TCheckBox(GetControl('SAUT5'));

  if (ComboRupture1<>nil) then ComboRupture1.OnChange:=ComboRupture1Change;
  if (ComboRupture2<>nil) then ComboRupture2.OnChange:=ComboRupture2Change;
  if (ComboRupture3<>nil) then ComboRupture3.OnChange:=ComboRupture3Change;
  if (ComboRupture4<>nil) then ComboRupture4.OnChange:=ComboRupture4Change;
  if (ComboRupture5<>nil) then ComboRupture5.OnChange:=ComboRupture5Change;

  // Prise en compte des valeurs par defaut
  if (ComboRupture5<>nil) then ComboRupture5Change(ComboRupture5);
  if (ComboRupture4<>nil) then ComboRupture4Change(ComboRupture4);
  if (ComboRupture3<>nil) then ComboRupture3Change(ComboRupture3);
  if (ComboRupture2<>nil) then ComboRupture2Change(ComboRupture2);
  if (ComboRupture1<>nil) then ComboRupture1Change(ComboRupture1);

  if (GetControl('XX_ORDERBY') is THValComboBox) then
  // dans le grand livre... le xx_orderby n'est pas un THValComboBox mais un THEdit invisible
  // donc inutile de gérer son caractère enabled ou pas dans ce cas...
  // On ne gère pas non plus les ORDERBY dont le tag=99
  begin
    ComboOrderBy := THValComboBox(GetControl('XX_ORDERBY'));
    if (ComboOrderBy <> nil) then
    begin
      if (ComboOrderBy.tag<>99) then
        sOrderByDefaut := ComboOrderBy.Value
      else
        ComboOrderBy:=nil;
    end;
  end;

  // Gestion des parties du code affaire invisibles à supprimer des combos
  NomsChampsAffaireString(sAff0, sAff1, sAff2, sAff3, sAff4);
  if (EditAff1<>nil) then
    if (EditAff1.Visible=false) then
    begin
      DeleteItemDansCombo(ComboRupture1, sAff1);
      DeleteItemDansCombo(ComboRupture2, sAff1);
      DeleteItemDansCombo(ComboRupture3, sAff1);
      DeleteItemDansCombo(ComboRupture4, sAff1);
      DeleteItemDansCombo(ComboRupture5, sAff1);
    end;

  if (EditAff2<>nil) then
    if (EditAff2.Visible=false) then
    begin
      DeleteItemDansCombo(ComboRupture1, sAff2);
      DeleteItemDansCombo(ComboRupture2, sAff2);
      DeleteItemDansCombo(ComboRupture3, sAff2);
      DeleteItemDansCombo(ComboRupture4, sAff2);
      DeleteItemDansCombo(ComboRupture5, sAff2);
    end;

  if (EditAff3<>nil) then
    if (EditAff3.Visible=false) then
    begin
      DeleteItemDansCombo(ComboRupture1, sAff3);
      DeleteItemDansCombo(ComboRupture2, sAff3);
      DeleteItemDansCombo(ComboRupture3, sAff3);
      DeleteItemDansCombo(ComboRupture4, sAff3);
      DeleteItemDansCombo(ComboRupture5, sAff3);
    end;

  // Gestion des champs ressource en fonction de la confidentialité
  NomsChampsRessource(Ressource1, Ressource2, LblRessource1, LblRessource2);
  if (Ressource1<>nil) then
  begin
    if Not SaisieActiviteManager then
    begin
      // Recherche de l'assistant à partir du user        
      //AFORessources:= TAFO_Ressources.Create;
      LOI := ListeObjectsInvisibles('RESS');
      try
        Ressource1.Text:=VH_GC.RessourceUser;
        //Ressource1.Text:=AFORessources.RessourceD1User(V_PGI.User);
        Ressource1.ElipsisButton := false;
        Ressource1.Enabled := false;
        if (LblRessource1<>nil) then begin LblRessource1.Caption := TraduitGA('Ressource'); end;
        if (Ressource2<>nil) then begin Ressource2.visible := false; Ressource2.Text:=Ressource1.Text; end;
        if (LblRessource2<>nil) then begin LblRessource2.visible:=false; end;
        // Rendre invisible la liste des champs fournis
        if (LOI<>nil) then
        for i:=0 to LOI.Count-1 do
            begin
            if (LOI.Objects[i] is TTabSheet) then
                TTabSheet(LOI.Objects[i]).TabVisible := false
            else
                TControl(LOI.Objects[i]).Visible := false;
            end;

      finally
        LOI.Free;
        //AFORessources.Free;
      end;
    end;
  end;

  // Gestion des champs invisibles suivant la confidentialité sur la valorisation
  if not AffichageValorisation then
  begin
    // Rendre invisible la liste des champs fournis
    LOI := ListeObjectsInvisibles('VALO');
    try
      if (LOI<>nil) then
        for i:=0 to LOI.Count-1 do
        begin
          // Les onglets
          if (LOI.Objects[i] is TTabSheet) then
            TTabSheet(LOI.Objects[i]).TabVisible := false
          else
            begin
              // les champs sur les onglets
              TControl(LOI.Objects[i]).Visible := false;

              // les champs dans la liste
              //if (Ecran is TFMul) then
              //if (TFMul(Ecran).FListe <> nil) then
              //For j:=0 to TFMul(Ecran).FListe.Columns.Count-1 do
              //    Begin
              //    If (TFMul(Ecran).FListe.Columns[j].FieldName = LOI[i])
              //        Then
              //        TFMul(Ecran).FListe.columns[j].visible:=False ;
              //    end;
            end;
        end;

    finally
      LOI.Free;
    end;
  end;

  //mcd 05/03/02 ajout de la condition plus sur type article paramétrable ...
  ThValMul:=THMultiValComboBox(GetControl('GA_TYPEARTICLE'));
  if (ThValMul =Nil) then ThValMul:=THMultiVALComboBox(GetControl('ACT_TYPEARTICLE'));
  if (ThValMul =Nil) then ThValMul:=THMultiVALComboBox(GetControl('GL_TYPEARTICLE'));
  if (ThValMul =Nil) then ThValMul:=THMultiVALComboBox(GetControl('EAC_TYPEARTICLE'));
  if (ThValMul =Nil) then ThValMul:=THMultiVALComboBox(GetControl('ATBXTYPEARTICLE'));
  if (ThValMul <>Nil) then
  begin
    ThValMul.Plus := PlusTypeArticle;
    if ThValMul.text='' then ThValMul.text:=PlusTypeArticleText;
  end;

  //
  //uniquement en line
  // Ecran.Caption := 'Liste des Chantiers';
     Ecran.Caption := 'Liste des Affaires';

  // mise à jour du titre du form suivant le contexte Scot/Tempo
  updatecaption(Ecran);

End;

procedure TOF_BASE_ETATS.NomsChampsRessource(var Res1, Res2:THEdit; var LblRes1, LblRes2:THLabel);
var
ResARS : THEdit;
begin
ResARS:=THEdit(GetControl('ARS_RESSOURCE'));
if ResARS<>nil then
    begin
    Res1 := ResARS;
    LblRes1:=THLabel(GetControl('TARS_RESSOURCE'));
    Res2:=THEdit(GetControl('ARS_RESSOURCE_'));
    LblRes2:=THLabel(GetControl('TARS_RESSOURCE_'));
    exit;
    end;

ResARS:=THEdit(GetControl('ACT_RESSOURCE'));
if ResARS<>nil then
    begin
    Res1 := ResARS;
    LblRes1:=THLabel(GetControl('TACT_RESSOURCE'));
    Res2:=THEdit(GetControl('ACT_RESSOURCE_'));
    LblRes2:=THLabel(GetControl('TACT_RESSOURCE_'));
    exit;
    end;

ResARS:=THEdit(GetControl('ATB_RESSOURCE'));
if ResARS<>nil then
    begin
    Res1 := ResARS;
    LblRes1:=THLabel(GetControl('TATB_RESSOURCE'));
    Res2:=THEdit(GetControl('ATB_RESSOURCE_'));
    LblRes2:=THLabel(GetControl('TATB_RESSOURCE_'));
    exit;
    end;

end;

procedure TOF_BASE_ETATS.NomsChampsAffaireString(var sAff0, sAff1, sAff2, sAff3, sAff4 : string);
begin
sAff0:= '';
sAff1:='AFF_AFFAIRE1';
sAff2:='AFF_AFFAIRE2';
sAff3:='AFF_AFFAIRE3';
sAff4:='';
end;

procedure TOF_BASE_ETATS.DeleteItemDansCombo(Combo:THValComboBox; Item:string);
var index:integer;
begin
if (Combo<>nil) then
   begin
   index := Combo.Values.IndexOf(Item);
   if (index<>-1) then
      begin
      Combo.Values.Delete(index);
      Combo.Items.Delete(index);
      end;
   end;
end;

procedure TOF_BASE_ETATS.CocherSiValeurDansCombo(ACocher:TCheckBox; sValeur:string; Combo:THValComboBox);
begin
if (ACocher<>nil) and (Combo<>nil) then
   if (Combo.Value=sValeur) then ACocher.Checked := true else ACocher.Checked := false;
end;

procedure TOF_BASE_ETATS.ComboRupture1Change(Sender: TObject);
begin
GereRuptureComboCheck(ComboRupture1, CheckTotRupt1,Saut1, ComboRupture2);
if (ComboOrderBy <> nil) then
    if (trim(ComboRupture1.value) <> '') then
        begin
        ComboOrderBy.Value := '';
        ComboOrderBy.Enabled := false;
        end
    else
        begin
        ComboOrderBy.Value := sOrderByDefaut;
        ComboOrderBy.Enabled := True;
        end;

end;

procedure TOF_BASE_ETATS.ComboRupture2Change(Sender: TObject);
begin
GereRuptureComboCheck(ComboRupture2, CheckTotRupt2,Saut2, ComboRupture3);
end;

procedure TOF_BASE_ETATS.ComboRupture3Change(Sender: TObject);
begin
GereRuptureComboCheck(ComboRupture3, CheckTotRupt3,Saut3, ComboRupture4);
end;

procedure TOF_BASE_ETATS.ComboRupture4Change(Sender: TObject);
begin
GereRuptureComboCheck(ComboRupture4, CheckTotRupt4,Saut4, ComboRupture5);
end;

procedure TOF_BASE_ETATS.ComboRupture5Change(Sender: TObject);
begin
GereRuptureComboCheck(ComboRupture5, CheckTotRupt5,Saut5, nil);
end;

procedure TOF_BASE_ETATS.GereRuptureComboCheck(ComboA:THValComboBox; CheckA,SautA:TCheckBox; ComboB:THValComboBox);
var
ValCombo:variant;
begin
if (ComboA.Value='') or (ComboA.Value=' ') then
   begin
   if (ComboB <> nil) then
      begin
      ComboB.Value:=' ';
      ComboB.Value:='';
      ComboB.Enabled:=false;
      end;

   if (CheckA<>nil) then
      begin
      CheckA.State:=cbUnChecked;
      CheckA.Enabled:=false;
      end;
   if (SautA<>nil) then
      begin
      SautA.State:=cbUnChecked;
      SautA.Enabled:=false;
      end;
   end
else
   begin
   if (ComboB <> nil) then
      begin
      ValCombo:=ComboB.Value;
      ComboB.Value:=' ';
      ComboB.Value:=ValCombo;
      ComboB.Enabled:=true;
      end;
   if (CheckA<>nil) then
      CheckA.Enabled:=true;
   if (SautA<>nil) then
      SautA.Enabled:=true;
   end;
end;


procedure TOF_BASE_ETATS.TableauObjectsInvisibles(Crit: string; var iNbChamps:integer; var tbChamps:PString);
begin
end;

function TOF_BASE_ETATS.ListeObjectsInvisibles(Crit:string):TStringList;
var
  i, iNbChamps  :integer;
  LOITempo      : TStringList;
  ControlTempo  : TControl;
  TbChpsInv     : PString;

begin
  iNbChamps := 0;
  LOITempo := TStringList.Create;
  TableauObjectsInvisibles(Crit, iNbChamps, TbChpsInv);

  // Scrute le tableau des champs invisibles pour remplir la liste demandée en sortie
  for i:=1 to iNbChamps do
  begin
    ControlTempo:=GetControl(string(TbChpsInv^));
    if (ControlTempo<>nil) then LOITempo.AddObject(string(TbChpsInv^), TObject(ControlTempo));
    Inc(TbChpsInv);
  end;

  Result := LOITempo;
end;


procedure TOF_BASE_ETATS.OnLoad;
begin
  inherited;

// Doublé dans le onargument (pour les QRS1, le onload n'est appelé qu'au lancement de l'état)
if (Ressource1<>nil) then
    begin
    if Not SaisieActiviteManager then
        begin
        Ressource1.Text:=VH_GC.RessourceUser;
        if (Ressource2<>nil) then begin Ressource2.Text:=Ressource1.Text; end;
        end;
    end;

end;

Initialization
registerclasses([TOF_BASE_ETATS]);
end.
