unit UTofEtatsAffaire;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,AGLInit,SaisUtil,HPanel,
      grids, Dicobtp, HSysMenu,HTB97,utofbaseetats,
{$IFDEF EAGLCLIENT}
   Maineagl,eMul,
{$ELSE}
   FE_Main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,  Mul,
{$ENDIF}
      URecupSQLModele,Paramsoc ,Ent1,
      HCtrls,HEnt1,HMsgBox,UTOF, FactUtil, EntGC,
      utob, Hqry;


Type
     TOF_ETATS_AFF = Class (TOF_BASE_ETATS)
      procedure OnArgument(stArgument : String ) ; override ;
      procedure OnClose ; override;
      procedure OnLoad ; override ;
      procedure OnUpdate ; override;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
      procedure ComboRupture1Change(Sender: TObject); override ;
      procedure ComboRupture2Change(Sender: TObject); override ;
      procedure ComboRupture3Change(Sender: TObject); override ;
      procedure ComboRupture4Change(Sender: TObject); override ;
      procedure ComboRupture5Change(Sender: TObject); override ;
      procedure TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString); override;
      procedure DateTraitementChange(Sender: TObject);
      procedure DateFinChange(Sender: TObject);
  private
      Bureau    : boolean;
      //
      procedure CheckedClick(Sender: TObject);
      procedure ControleChamp(Champ, Valeur: String);

  public
      TMoisCloture   :  THEdit;
      TMoisCloture_  :  THEdit;
      TTMoisCloture  :  THLabel;
      TTMoisCloture_ :  THLabel;
      DateDebExer    :  THEdit;
      DateDebExer_   :  THEdit;
      TDateDebExer   :  THLabel;
      TDateDebExer_  :  THLabel;
      DateFinExer    :  THEdit;
      DateFinExer_   :  THEdit;
      TDateFinExer   :  THLabel;
      TDateFinExer_  :  THLabel;
      TabSheetChampParam : TTabSheet;
      DateTraitement :  THEdit;
      DateFinGener ,DateFinGener2  :  THEdit;
      ChangeDateFin: Boolean;      // mcd 08/10/02
      //
      MEP_FApercu  : TCheckBox;
      MEP_FReduire : TCHeckBox;
      MEP_FCouleur : TCHeckBox;
      MEP_Titre 	 : TCHeckBox;
      MEP_FListe	 : TCHeckBox;
      //
     END ;

const
NbTbChRessInv = 13 ;
TbChampsRessInvisibles : array[1..NbTbChRessInv] of string 	= (
          {1}        'ARS_LIBELLE',
          {2}        'TARS_LIBELLE',
          {3}        'ARS_LIBELLE_',
          {4}        'TARS_LIBELLE_',
          {5}        'ARS_FONCTION1',
          {6}        'TARS_FONCTION1',
          {7}        'ARS_FONCTION1_',
          {8}        'TARS_FONCTION1_',
          {9}        'ARS_TYPERESSOURCE',
          {10}       'TARS_TYPERESSOURCE',
          {11}       'STATISTIQUES',
          {12}       'COMPLEMENT',
          {13}       'RUPTURE'
          );

NbTbChValoInv = 10 ;
TbChampsValoInvisibles : array[1..NbTbChValoInv] of string 	= (
          {1}        'ARS_PVHT',
          {2}        'TARS_PVHT',
          {3}        'TARS_PVHT_',
          {4}        'ARS_PVHT_',
          {5}        'TARS_TAUXREVIENTUN',
          {6}        'ARS_TAUXREVIENTUN',
          {7}        'TARS_TAUXREVIENTUN_',
          {8}        'ARS_TAUXREVIENTUN_',
          {9}        'TXX_VARIABLE1',
          {10}       'XX_VARIABLE1'
          );


Procedure AFLanceFiche_EditAdresseClient;
Procedure AFLanceFiche_EditAdresseRess;
Procedure AFLanceFiche_EtatCommission;
Procedure AFLanceFiche_EditPrepFact;
Procedure AFLanceFiche_EtatPrevCa(Argument : string);
Procedure AFLanceFiche_EtatsAffaire;
Procedure AFLanceFiche_EtatsClient;
Procedure AFLanceFiche_EtatsPrestation;
Procedure AFLanceFiche_EtatsRess;
Procedure AFLanceFiche_EtiquetteClient;
Procedure AFLanceFiche_FicheAffaire;
Procedure AFLanceFiche_FicheClient;
Procedure AFLanceFiche_FicheRessource;
Procedure AFLanceFiche_EditStatAffaire;
Procedure AFLanceFiche_EditStatClient;
Procedure AFLanceFiche_EditStatPrest;
Procedure AFLanceFiche_EditStatRess;


implementation


procedure TOF_ETATS_AFF.OnClose ;
begin
Inherited;
If Ecran.name ='ETIQUETTECLIENT' then
   ExecuteSQL('DELETE FROM GCTMPETQCLI WHERE GZK_UTILISATEUR = "'+V_PGI.USer+'"');
end;


procedure TOF_ETATS_AFF.OnUpdate ;
var stWhere, stInsert : string;
    QEtiq : TQuery;
    iInd2, nbEtiq, Cptr : integer;
begin
Inherited;
If Ecran.name ='ETIQUETTECLIENT' then
  begin   // repris de UtofEtiqCli : fait pour permettre d'imprimer X etiquette pour un même client
  ExecuteSQL('DELETE FROM GCTMPETQCLI WHERE GZK_UTILISATEUR = "'+V_PGI.USer+'"');
  SetControlText('XX_WHERE_USER','');
  stWhere := RecupWhereCritere(TPageControl(GetControl('PAGES')));
  SetControlText('XX_WHERE_USER','GZK_UTILISATEUR="'+V_PGI.USer+'"');
  QEtiq := OpenSQL('SELECT T_TIERS FROM TIERS LEFT JOIN TIERSCOMPL ON LIEN(TIERS,TIERSCOMPL) ' + stWhere, true);
  Cptr:=0;
  While not QEtiq.Eof do  // repris UtofEtiqCli
    begin
    nbEtiq := StrToInt(GetControlText('NBEXEMPLAIRE'));
    for iInd2 := 0 to nbEtiq - 1 do
      begin
      stInsert := 'Insert into GCTMPETQCLI (GZK_UTILISATEUR, GZK_COMPTEUR, GZK_TIERS,GZK_NBETIQ) ' +
            'SELECT "' + V_PGI.User + '" as GZK_UTILISATEUR, ' 
            + IntToStr(Cptr+1) + ' as GZK_COMPTEUR, ' +
            ' "'+QEtiq.FindField('T_TIERS').AsString+'" as GZK_TIERS,'
            + IntToStr(NbEtiq)+ ' as GZK_ETIQ';
      ExecuteSQL(stInsert);
      Inc(Cptr);
      end;
    QEtiq.Next;
    end;
  Ferme(QEtiq);
  end;
end;


procedure TOF_ETATS_AFF.OnLoad;
Var Combo: ThValCOmboBox;
    stsql,smodele : string;
begin
  inherited;
  // mcd 30/07/2002 etat prev CA, ajout tri affaire obligatoire si pas choisi

If (Ecran).name='ETATPREVCA' then begin
    smodele :=  THValComboBox(GetControl('FETAT')).Value;
    stSQL:=RecupSQLModele('E', 'ACA',smodele ,'','','',' ');
    // gm 24/03/03
    // modif  états spécif d'ALGOE, car dans un état  specif il n'y a pas de jointure sur table FACTAFF
    // donc le where fait planter
    // modif provixsoire je regarde si dans la requete on trouve le mot " LIGNE "
      if pos('FACTAFF',stSql) > 0 then
        SetControlText('XX_WHERE', 'AFA_DATEECHE>="'+UsDAteTime(StrToDaTe(GetCOntrolText('DATETRAITEMENT')))+'"');

   Combo := THValComboBox(GetControl('XX_ORDERBY1'))  ;
   SetCOntrolText('XX_ORDERBY',Combo.value);
   If COmbo.value <>''   then
      begin
      if pos('AFF_AFFAIRE',combo.value)=0
           then  begin
           SetCOntrolText('XX_ORDERBY',Combo.value + ',AFF_AFFAIRE');
           end;
      end
     Else    SetCOntrolText('XX_ORDERBY','AFF_TIERS,AFF_AFFAIRE');
   end;

end;


procedure TOF_ETATS_AFF.OnArgument(stArgument : String );
var Combo : ThValComboBox;
 check : TCheckBox;
 CC : THValComboBox;
    CCE       : Thedit;
 StRupture : String;
    Critere   : string;
    ValMul    : string;
    ChampMul  : string;
    x         : integer;
begin

   Bureau := False;
   //Critere:=(Trim(ReadTokenSt(stArgument)));
   //
  Repeat
    Critere:=uppercase(ReadTokenSt(stArgument)) ;
    valMul := '';
       if Critere<>'' then
    begin
      x:=pos('=',Critere);
          if x<>0 then
             begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
             end
      else
        ChampMul := Critere;
      ControleChamp(ChampMul, ValMul);
             end;
  until Critere='';

  // Suppression de l'onglet Champs paramétrables en Gestion d'affaire
  TabSheetChampParam := TTabSheet(GetControl('COL_VARIABLE'));

  // mcd 25/02/2003
  Check := TCheckBox (GetControl('ISPLANNING'));
  if Check <> Nil then
   begin
   if VH_GC.GAPlanningSeria then
      SetcontrolChecked('ISPLANNING',true)
   else
      SetcontrolChecked('ISPLANNING',False);
   end;

  // Suppression des champs mois de cloture en Gestion d'Affaire
  TMoisCloture := THEdit(GetControl('T_MOISCLOTURE'));
  TMoisCloture_ := THEdit(GetControl('T_MOISCLOTURE_'));
  TTMoisCloture := THLabel(GetControl('TT_MOISCLOTURE'));
  TTMoisCloture_ := THLabel(GetControl('TT_MOISCLOTURE_'));

  if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   if (TMoisCloture<>nil) then TMoisCloture.visible:=false;
   if (TMoisCloture_<>nil) then TMoisCloture_.visible:=false;
   if (TTMoisCloture<>nil) then TTMoisCloture.visible:=false;
   if (TTMoisCloture_<>nil) then TTMoisCloture_.visible:=false;
   if (TabSheetChampParam<>nil) then TabSheetChampParam.TabVisible:=false;
   end;

  // Suppression des champs date debut d'exercice en Gestion d'Affaire
  DateDebExer := THEdit(GetControl('AFF_DATEDEBEXER'));
  DateDebExer_ := THEdit(GetControl('AFF_DATEDEBEXER_'));
  TDateDebExer := THLabel(GetControl('TAFF_DATEDEBEXER'));
  TDateDebExer_ := THLabel(GetControl('TAFF_DATEDEBEXER_'));

  if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   if (DateDebExer<>nil) then DateDebExer.visible:=false;
   if (DateDebExer_<>nil) then DateDebExer_.visible:=false;
   if (TDateDebExer<>nil) then TDateDebExer.visible:=false;
   if (TDateDebExer_<>nil) then TDateDebExer_.visible:=false;
   end;

  // Suppression des champs date fin d'exercice en Gestion d'Affaire
  DateFinExer := THEdit(GetControl('AFF_DATEFINEXER'));
  DateFinExer_ := THEdit(GetControl('AFF_DATEFINEXER_'));
  TDateFinExer := THLabel(GetControl('TAFF_DATEFINEXER'));
  TDateFinExer_ := THLabel(GetControl('TAFF_DATEFINEXER_'));

  if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   if (DateFinExer<>nil) then DateFinExer.visible:=false;
   if (DateFinExer_<>nil) then DateFinExer_.visible:=false;
   if (TDateFinExer<>nil) then TDateFinExer.visible:=false;
   if (TDateFinExer_<>nil) then TDateFinExer_.visible:=false;
   end;

Inherited;

  //BchangeTiers := BBCHangeTiers;

   // On traite la suppression de T_MOISCLOTURE dans les combo associées aux ruptures
   if Not (ctxScot in V_PGI.PGIContexte) then
      begin
      DeleteItemDansCombo(ComboRupture1, 'T_MOISCLOTURE');
      DeleteItemDansCombo(ComboRupture2, 'T_MOISCLOTURE');
      DeleteItemDansCombo(ComboRupture3, 'T_MOISCLOTURE');
      DeleteItemDansCombo(ComboRupture4, 'T_MOISCLOTURE');
      DeleteItemDansCombo(ComboRupture5, 'T_MOISCLOTURE');
      end;

   // Gestion du cochage des champs RuptAff si la combo XX_RUPTURE = GP_AFFAIRE
   CocherSiValeurDansCombo(RuptAff1, 'GP_AFFAIRE', ComboRupture1);
   CocherSiValeurDansCombo(RuptAff2, 'GP_AFFAIRE', ComboRupture2);
   CocherSiValeurDansCombo(RuptAff3, 'GP_AFFAIRE', ComboRupture3);
   CocherSiValeurDansCombo(RuptAff4, 'GP_AFFAIRE', ComboRupture4);
   CocherSiValeurDansCombo(RuptAff5, 'GP_AFFAIRE', ComboRupture5);

   // Gestion de la synchronisation entre le champ DateTraitement et DateFinGener
   DateTraitement := THEdit(GetControl('DATETRAITEMENT'));
   DateFinGener := THEdit(GetControl('AFF_DATEFINGENER'));
   DateFinGener2 := THEdit(GetControl('AFF_DATEFINGENER_'));
   //
   if (DateTraitement<>nil) and (DateFinGener<>nil) then
       begin
       DateTraitementChange(nil);
       //DateTraitement.OnChange:=DateTraitementChange; mcd 18/03/2002 pour ne le faire qu'a la sortie du champ
       DateTraitement.OnExit:=DateTraitementChange;
       DateFinGener2.OnExit:=DateFinChange;  //mcd 08/10/02
       ChangeDateFin:=False;
       end;
   //
  if (ctxScot in V_PGI.PGIContexte) then
  begin
    if (Ecran).name='ETATSCLIENT' then
    begin  //mcd 08/03/02
         Combo := ThValCombobox(Getcontrol ('XX_VARIABLE1'));
         Combo.plus:='And (CO_CODE like "T%" or CO_CODE ="CB4")';
         Combo.value:='T_SECTEUR';
         Combo := ThValCombobox(Getcontrol ('XX_VARIABLE2'));
         Combo.plus:='And (CO_CODE like "T%" or CO_CODE ="CB4")';
         Combo.value:='T_MOISCLOTURE';
         Combo := ThValCombobox(Getcontrol ('XX_VARIABLE5'));
         Combo.plus:='And (CO_CODE like "T%" or CO_CODE ="CB4")';
         Combo.value:='T_MODEREGLE';
         Combo := ThValCombobox(Getcontrol ('XX_VARIABLE4'));
         Combo.plus:='And (CO_CODE like "T%" or CO_CODE ="CB4")';
         Combo.value:='T_REGIMETVA';
         end;
      end;

   // ajout mcd 06/05/2003
  If GetControl('Afa_GenerAuto') <> Nil then
  begin
         // si modif voir Onargument de UtofAfPrepFact
     SetControlProperty('AFA_GENERAUTO','Plus','GA');
         // ohligation de remplir texte pour ne pas prendre MAN si TOUT
     SetControlTExt('AFA_GENERAUTO','CON;FOR;MAN;POT;POU;ACT');
    if ctxScot in V_PGI.PGIContexte then
    begin
          SetControlProperty('AFA_GENERAUTO','Plus','GA" AND CO_CODE<>"MAN" AND CO_CODE<>"CON');
          SetControlTExt('AFA_GENERAUTO','FOR;POT;POU;ACT');
          end;
     end;
  //
  If GetControl('Afa_LIQUIDATIVE') <> Nil then
  begin
    If Not (GetParamSoc('SO_AFGERELIQUIDE')) then
     begin
     SetControlVisible('AFA_LIQUIDATIVE',False);
     SetControlChecked ('AFA_LIQUIDATIVE',False);
     end;
    end;
  //
   If Bureau then
    begin   //mcd 09/10/03
    SetControlProperty ('BAGRANDIR','down',true);
    //TToolbarButton97(GetControl('BAGRANDIR')).Onclick(self);
    TpageControl(GEtControl('Pages')).visible :=False;
    TToolbarButton97(GetControl('BVALIDER')).Onclick(self);
    end;
   // gestion Etablissement (BTP)
  if GetControl('AFF_ETABLISSEMENT') IS THValComboBox then CC:=THValComboBox(GetControl('AFF_ETABLISSEMENT'));
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  if GetControl('AFF_DOMAINE') IS THValComboBox then CC:=THValComboBox(GetControl('AFF_DOMAINE'));
  if CC<>Nil then PositionneDomaineUser(CC) ;

// --
//uniquement en line
{*

   //
   MEP_FApercu  := TCheckBox(GetControl('FAPERCU1'));
   MEP_FReduire := TCheckBox(GetControl('FREDUIRE1'));
   MEP_FCouleur := TCheckBox(GetControl('FCOULEUR1'));
   MEP_Titre 	 	:= TCheckBox(GetControl('TITRE1'));
   MEP_FListe	 	:= TCheckBox(GetControl('FLISTE1'));
   //
   MEP_FApercu.OnClick := CheckedClick;
   MEP_FReduire.OnClick := CheckedClick;
   MEP_FCouleur.OnClick := CheckedClick;
   MEP_Titre.OnClick := CheckedClick;
   MEP_FListe.OnClick := CheckedClick;
   //
   //Suppression de l'onglet mise en Page
   TTabSheet(GetControl('Option')).TabVisible := False;
   //
   if THValComboBox(GetControl('AFF_ETABLISSEMENT'))<>Nil then THValComboBox(GetControl('AFF_ETABLISSEMENT')).visible := false;
   if THLabel(GetControl('TAFF_ETABLISSEMENT'))<>Nil then THLabel(GetControl('TAFF_ETABLISSEMENT')).visible := false;
   if THValComboBox(GetControl('FETAT'))<>Nil then THValComboBox(GetControl('FETAT')).visible := false;
   if THLabel(GetControl('TETAT'))<>Nil then THLabel(GetControl('TETAT')).visible := false;
   if THEDIT(GetControl('AFF_RESPONSABLE'))<>Nil then THEDIT(GetControl('AFF_RESPONSABLE')).visible := false;
   if THLabel(GetControl('TAFF_RESPONSABLE'))<>Nil then THLabel(GetControl('TAFF_RESPONSABLE')).visible := false;
   if THMultiValComboBox(GetControl('AFF_GENERAUTO'))<>Nil then THMultiValComboBox(GetControl('AFF_GENERAUTO')).Plus := 'BTP" AND CO_CODE <> "DAC';
   //
   Ecran.Caption := 'Liste des Chantiers';
   //
   SetControlText('TAFF_AFFAIRE1', 'Chantier');
   SetControlText('TAFF_RESPONSABLE', 'Resp. Chantier');
   SetControlText('TAFF_ETATAFFAIRE', 'Etat du Chantier');
   //
   THValComboBox(GetControl('XX_RUPTURE1')).Items := THValComboBox(GetControl('XX_RUPTURE_S1')).Items ;
   THValComboBox(GetControl('XX_RUPTURE1')).Values := THValComboBox(GetControl('XX_RUPTURE_S1')).Values ;
   //On remplace OldText 'Affaire' par NewText 'Chantier'
   for X := 0 to THValComboBox(GetControl('XX_RUPTURE1')).items.Count - 1 Do
       Begin
       StRupture := THValComboBox(GetControl('XX_RUPTURE1')).items[X];
       StRupture := StringReplace(StRupture,'affaire', 'Chantier',[rfReplaceAll]);
       THValComboBox(GetControl('XX_RUPTURE1')).items[X] := Strupture;
       end;

	 THValComboBox(GetControl('XX_RUPTURE2')).Items := THValComboBox(GetControl('XX_RUPTURE1')).Items ;
	 THValComboBox(GetControl('XX_RUPTURE2')).Values := THValComboBox(GetControl('XX_RUPTURE1')).Values ;

   THValComboBox(GetControl('XX_ORDERBY')).Items := THValComboBox(GetControl('XX_ORDERBY_S1')).Items ;
   THValComboBox(GetControl('XX_ORDERBY')).Values := THValComboBox(GetControl('XX_ORDERBY_S1')).Values ;

   //On remplace OldText 'Affaire' par NewText 'Chantier'
   for X := 0 to THValComboBox(GetControl('XX_ORDERBY')).items.Count - 1 Do
       Begin
       StRupture := THValComboBox(GetControl('XX_ORDERBY')).items[X];
       StRupture := StringReplace(StRupture,'affaire', 'Chantier',[rfReplaceAll]);
       THValComboBox(GetControl('XX_ORDERBY')).items[X] := Strupture;
       end;

*}
	// gestion Chargé affaire (BTP)
	CCE:=ThEdit(GetControl('AFF_RESPONSABLE')) ;
	if CCE<>Nil then
  begin
  	PositionneResponsableUser(CCE) ;
	end;

End;

Procedure TOF_ETATS_AFF.ControleChamp(Champ : String;Valeur : String);
begin

  if Champ = 'TIERS' then
  begin
    SetControlText('AFF_TIERS',Valeur);
    SetControlEnabled ('AFF_TIERS',False);
    SetControlText('AFF_TIERS_',Valeur);
    SetControlEnabled ('AFF_TIERS_',False);
    BchangeTiers :=False;
  end
  else if (Champ='BUREAU') and (Ecran.name='ETATPREVCA') then
  begin   //mcd 09/10/03 pour acces état depuis bureau.. force des valeurs
    SetControlText ('XX_RUPTURE1','AFF_ETABLISSEMENT');
    SetControlChecked ('TOTAL1',True);
    SetControlChecked ('TITRE',false);
    Bureau :=True;
  end
  Else If Champ = 'STATUT' then
  Begin
    if Valeur = 'APP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Appel');
      SetControlText('TAFF_AFFAIRE', 'Appel');
    end
    Else if valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="I"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Contrat');
      SetControlText('TAFF_AFFAIRE1', 'Contrat');
    end
    Else if valeur = 'AFF' then
    Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then
      Begin
        SetControlText('AFF_AFFAIRE0', 'A');
        SetControltext('XX_WHERE',' AND AFF_AFFAIRE in ("A", "")');
      end
      else if assigned(GetControl('AFFAIRE0')) then
      begin
        SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) in ("A", "")');
        SetControlText('AFFAIRE0', 'A');
      end;
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Chantier');
      SetControlText('TAFF_AFFAIRE1', 'Chantier');
    end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControltext('XX_WHERE',' AND AFF_AFFAIRE0 IN ("A", "W")')
      else if assigned(GetControl('AFFAIRE0')) then SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) in ("A", "W")');
    end
    Else if valeur = 'PRO' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Appel d''offre');
      SetControlText('TAFF_AFFAIRE1', 'Appel d''offre');
    end
    else
    Begin
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Affaire');
      SetControlText('TAFF_AFFAIRE1', 'Affaire');
    end
  end;

end;

Procedure TOF_ETATS_AFF.CheckedClick(Sender: TObject);
Begin

	TCheckBox(GetControl('FAPERCU')).Checked 	:= MEP_FApercu.Checked;
  TCheckBox(GetControl('FREDUIRE')).Checked := MEP_FReduire.Checked;
  TCheckBox(GetControl('FCOULEUR')).Checked := MEP_FCouleur.Checked;
  TCheckBox(GetControl('TITRE')).Checked 		:= MEP_Titre.Checked;
	TCheckBox(GetControl('FLISTE')).Checked 	:= MEP_Fliste.Checked;

End;

procedure TOF_ETATS_AFF.DateTraitementChange(Sender: TObject);
var date : TdateTime;
begin
     // mcd 18/03/2002 : ajout calcul date fin, sinon initialisé à la date de debut
DateFinGener.Text := DateTraitement.Text;
 // mcd 08/10/02 si date fin chnager, on en recalcule pas sauf si >
if Not(ChangeDateFin) or (StrToDAte(DateFinGener.Text) >StrToDAte(DAteFinGener2.text)) then begin
   Date := Plusdate(StrToDate(DateTraitement.text),15,'M');
   Date := Plusdate (Date,-1,'J');
   DAteFinGener2.text := DateToStr(Date);
   end;
end;

procedure TOF_ETATS_AFF.DateFinChange(Sender: TObject);
begin
  // mcd08/10/02 si la date de fin a été changé, on le note
  ChangeDAteFin := True;
end;

procedure TOF_ETATS_AFF.ComboRupture1Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff1, 'GP_AFFAIRE', ComboRupture1);
end;

procedure TOF_ETATS_AFF.ComboRupture2Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff2, 'GP_AFFAIRE', ComboRupture2);
end;

procedure TOF_ETATS_AFF.ComboRupture3Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff3, 'GP_AFFAIRE', ComboRupture3);
end;

procedure TOF_ETATS_AFF.ComboRupture4Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff4, 'GP_AFFAIRE', ComboRupture4);
end;

procedure TOF_ETATS_AFF.ComboRupture5Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff5, 'GP_AFFAIRE', ComboRupture5);
end;

procedure TOF_ETATS_AFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
// mcd 19/09/01 Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
// mcd 19/09/01 ne pas alimenter pour OK sélection aff Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

procedure TOF_ETATS_AFF.TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);
begin
if (Crit='RESS') then
    begin
    iNbChamps := NbTbChRessInv;
    tbChamps := @TbChampsRessInvisibles;
    end;

if (Crit='VALO') then
    begin
    iNbChamps := NbTbChValoInv;
    tbChamps := @TbChampsValoInvisibles;
    end;
end;


Procedure AFLanceFiche_EditAdresseClient;
begin
AGLLanceFiche ('AFF','ADRESSECLIENT','','','');
end;
Procedure AFLanceFiche_EditAdresseRess;
begin
AGLLanceFiche ('AFF','ADRESSERESSOURCE','','','');
end;
Procedure AFLanceFiche_EtatCommission;
begin
AGLLanceFiche ('AFF','ETATCOMMISSION','','','');
end;
Procedure AFLanceFiche_EtatPrevCa (Argument : string);
begin
AGLLanceFiche ('AFF','ETATPREVCA','','',Argument);
end;
Procedure AFLanceFiche_EDitPrepFact;
begin
AGLLanceFiche ('AFF','AFEDITPREPFACT','','','');
end;
Procedure AFLanceFiche_EtatsAffaire;
begin
AGLLanceFiche ('AFF','ETATSAFFAIRE','','','');
end;
Procedure AFLanceFiche_EtatsClient;
begin
AGLLanceFiche ('AFF','ETATSCLIENT','','','');
end;
Procedure AFLanceFiche_EtatsPrestation;
begin
AGLLanceFiche ('AFF','ETATSPRESTATION','','','');
end;
Procedure AFLanceFiche_EtatsRESS;
begin
AGLLanceFiche ('AFF','ETATSRESSOURCE','','','');
end;

Procedure AFLanceFiche_EtiquetteClient;
begin
AGLLanceFiche ('AFF','ETIQUETTECLIENT','','','');
end;
Procedure AFLanceFiche_FicheAffaire;
begin
AGLLanceFiche ('AFF','FICHEAFFAIRE','','','');
end;
Procedure AFLanceFiche_FicheClient;
begin
AGLLanceFiche ('AFF','FICHECLIENT','','','');
end;
Procedure AFLanceFiche_FicheRessource;
begin
AGLLanceFiche ('AFF','FICHERESSOURCE','','','');
end;
Procedure AFLanceFiche_EditStatAffaire;
begin
AGLLanceFiche ('AFF','STATAFFAIRE','','','');
end;
Procedure AFLanceFiche_EditStatClient;
begin
AGLLanceFiche ('AFF','STATCLIENT','','','');
end;
Procedure AFLanceFiche_EditStatPrest;
begin
AGLLanceFiche ('AFF','STATPRESTATION','','','');
end;
Procedure AFLanceFiche_EditStatRess;
begin
AGLLanceFiche ('AFF','STATRESSOURCE','','','');
end;

Initialization
registerclasses([TOF_ETATS_AFF]);
end.
