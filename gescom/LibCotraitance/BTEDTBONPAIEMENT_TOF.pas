{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/06/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTEDTBONPAIEMENT ()
Mots clefs ... : TOF;BTEDTBONPAIEMENT
*****************************************************************}
Unit BTEDTBONPAIEMENT_TOF;

Interface

Uses  StdCtrls,
      Controls,
      Classes,
      forms,
      sysutils,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOF,
      UTOB,
{$IFDEF EAGLCLIENT}
      MaineAGL,
      emul,
{$ELSE}
    {$IFNDEF ERADIO}
      Fe_Main,
      uPDFBatch,
    {$ENDIF !ERADIO}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      HDB,
      Mul,
      UtilGc,
{$ENDIF}
			hPDFPrev,
      hPDFViewer,
      EntGC,
      utofAfBaseCodeAffaire,
      HTB97,
      Hpanel,
      uEntCommun,
      UtilTOBPiece,
      uRecupSQLModele,
      UtilsEtat,
      paramsoc,
      Ent1,
      UCotraitance;

Type

  TOF_BTEDTBONPAIEMENT = Class (TOF_AFBASECODEAFFAIRE)

  public
    procedure OnArgument(stArgument : String ) ; override ;
    procedure OnClose; override;
    procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;

  private
    TobPIECE      : TOB;
    TobAffaire    : TOB;
    TobPiecTrait  : TOB;
    TobCalcul     : TOB;
    TobMandataire : TOB;
    TobSousTraite : TOB;
    TobTiersFact  : TOB;
    TobSociete    : TOB;
    TOBPieceRG    : TOB;
    TOBBasesRg    : TOB;
    TOBEC         : TOB;
    TOBES         : TOB;
    TOBEditions   : TOB;
    TOBRecap      : TOB;
    //
    Idef          : Integer;
    //
    Fgestion      : Tmodegestion;
    //
    Ok_Recap      : Boolean;
    Ok_Multiple   : Boolean;
    Ok_LastEdit   : Boolean;
    Ok_EdtSsTrait : Boolean;
    //
    BOuvrir       : TToolbarButton97;
    //
    BParamEtat    : TToolBarButton97;
    BParamEtat1   : TToolBarButton97;
    BParamEtat2   : TToolBarButton97;
    BParamEtat3   : TToolBarButton97;
    //
    ChkApercu     : TCheckBox;
    ChkReduire    : TCheckBox;
    ChkCouleur    : TCheckBox;
    ChkEditjust   : TCheckBox;
    //
    Fliste	      : THDbGrid;
    //
    GRB_FRAIS     : TGroupBox;

    // Cotraitant Eclatement
    FETAT         : THValComboBox;
    TEtat         : ThLabel;

    OptionEdition : TOptionEdition;
    TheType       : String;
    TheNature     : String;
    TheTitre      : String;
    TheModele     : String;

    // Cotraitant Paiement Direct
    FETAT1        : THValComboBox;
    TEtat1        : ThLabel;
    OptionEdition1: TOptionEdition;
    TheType1      : String;
    TheNature1    : String;
    TheTitre1     : String;
    TheModele1    : String;

    // Sous-Traitant
    FETAT2        : THValComboBox;
    OptionEdition2: TOptionEdition;
    TheType2      : String;
    TheNature2    : String;
    TheTitre2     : String;
    TheModele2    : String;
    // Frais
    FETAT3        : THValComboBox;
    OptionEdition3: TOptionEdition;
    TheType3      : String;
    TheNature3    : String;
    TheTitre3     : String;
    TheModele3    : String;
    //
    Cledoc        : r_cledoc;
    //
    procedure AddlesChamps (TOBEdt,TOBSrc : TOB);
    procedure AjouteLesParamsEdt (TobEDT: TOB; OptEdt : TOptionEdition);
    Procedure ChargeEnteteEdition(TobEdition: TOB; TypeInter : String);
    procedure ChargeEtatEdition;
    procedure ChargeEtatEdtEclatement;
    procedure ChargeEtatEdtDirecteCoTraite;
    procedure ChargeEtatEdtFrais;
    procedure ChargeEtatEdtSsTraitant;
    procedure ChargeIntervenantEdition(TobMere, TobCalc: TOB);
    procedure Chargementlignefrais;
    procedure ChargementMandataire(TOBPiece : TOB);
    procedure ChargeMandataireEdition(TobLigEdt, Tobcalc : TOB);
    procedure ChargeRecapEdition(TobRecap, TobCalc: TOB);
    procedure ChargementSociete;
    procedure ChargementTiersFacture(TiersFacture: string);
    procedure ConstitueTobCalcul;
    procedure ChargelesTobsPrincipales(Cledoc: R_Cledoc);
    procedure ControleChamp(Champ, Valeur: String);
    procedure ConstitueEditions;
    procedure CreateTobs;
    procedure CreationLigneMandataire;
    procedure EditeTout;
    procedure EditionDirecte(TobEdition, TobCalcul: TOB);
    procedure EditionRegroupement(TobEdition, TobCalcul: TOB; TotalMan: Double);
    procedure GenereRecap;
    procedure GetObjects;
    procedure InitScreenObject;
    procedure OnChangeEtat(Sender : TObject);
    procedure OnClickApercu(Sender: Tobject);
    procedure OnClickReduire(Sender: Tobject);
    procedure ParamEtat(Sender : TOBJect);
    procedure PrepareEditions(TOBEC, TOBES: TOB);
    procedure PrepareTOBEDit (TOBEc,TobedtC,TOBEdtS : TOB);
    procedure PrepareTOBRecap(TOBEc, TOBedtR: TOB);
    Procedure ReleaseTob;
    procedure SetScreenEvents;
    Procedure TraitelesLignes(Sender : TOBJect);
    procedure VisibleScreenObject;
    procedure OnClickEditJust(Sender: Tobject);
    procedure AddlesChampsSupMandataire(TOBmandataire: TOB);


  end;

implementation

uses  TntStdCtrls,
      BTPUtil,
      TiersUtil,
      Factutil,
      FactRG;

procedure TOF_BTEDTBONPAIEMENT.OnArgument(stArgument : String );
var CC 		    : THValComboBox ;
    Critere   : string;
    ValMul    : string;
    ChampMul  : string;
    x         : integer;
Begin
  inherited;

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

  GetObjects;
  SetScreenEvents;
  VisibleScreenObject;
  InitScreenObject;

  Ok_EdtSsTrait := false;

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  //chargement de l'état en fonction du type de paiement : si direct Attestation, si regroupement eclatement
  ChargeEtatEdition;

  VisibleScreenObject;

end;

Procedure TOF_BTEDTBONPAIEMENT.GetObjects;
begin

  if Assigned(GetControl('GRB_FRAIS'))   then GRB_FRAIS   := TGroupBox(ecran.FindComponent('GRB_FRAIS'));

  if Assigned(GetControl('BParamEtat'))  then BParamEtat  := TToolbarButton97(ecran.FindComponent('BParamEtat'));
  if Assigned(GetControl('BParamEtat1')) then BParamEtat1 := TToolbarButton97(ecran.FindComponent('BParamEtat1'));
  if Assigned(GetControl('BParamEtat2')) then BParamEtat2 := TToolbarButton97(ecran.FindComponent('BParamEtat2'));
  if Assigned(GetControl('BParamEtat3')) then BParamEtat3 := TToolbarButton97(ecran.FindComponent('BParamEtat3'));

  if Assigned(GetControl('Bouvrir'))   then BOuvrir     := TToolbarButton97(ecran.FindComponent('BOuvrir'));
  if Assigned(GetControl('Fliste'))    then Fliste      := THDbGrid(ecran.FindComponent('Fliste'));
  if Assigned(GetControl('fApercu'))   then ChkApercu   := TCheckBox(Ecran.FindComponent('fApercu'));
  if Assigned(GetControl('FReduire'))  then ChkReduire  := TCheckBox(Ecran.FindComponent('fReduire'));
  if Assigned(GetControl('fCouleur'))  then ChkCouleur  := TCheckBox(Ecran.FindComponent('fCouleur'));
  if Assigned(GetControl('fEDITJUST')) then ChkEditJust := TCheckBox(Ecran.FindComponent('fEDITJUST'));

  if Assigned(GetControl('fEtat'))     then FEtat   := ThValComboBox(ecran.FindComponent('FEtat'));
  if Assigned(GetControl('fEtat1'))    then FEtat1  := ThValComboBox(ecran.FindComponent('FEtat1'));
  if Assigned(GetControl('fEtat2'))    then FEtat2  := ThValComboBox(ecran.FindComponent('FEtat2'));
  if Assigned(GetControl('fEtat3'))    then FEtat3  := ThValComboBox(ecran.FindComponent('FEtat3'));
  if assigned(Getcontrol('TEtat'))     then TEtat   := ThLabel(ecran.FindComponent('TEtat'));
  if assigned(Getcontrol('TEtat1'))    then TEtat1  := ThLabel(ecran.FindComponent('TEtat1'));

end;

procedure TOF_BTEDTBONPAIEMENT.SetScreenEvents;
begin

  if assigned(BOuvrir)      then BOuvrir.OnClick      := TraiteLesLignes;

  if assigned(ChkApercu)    then ChkApercu.OnClick    := OnClickApercu;
  if assigned(ChkReduire)   then ChkReduire.OnClick   := OnClickReduire;
  if assigned(ChkEditJust)  then ChkEditJust.OnClick  := OnClickEditJust;

  if Assigned(BParamEtat)   then BParamEtat.OnClick   := ParamEtat;
  if Assigned(BParamEtat1)  then BParamEtat1.OnClick  := ParamEtat;
  if Assigned(BParamEtat2)  then BParamEtat2.OnClick  := ParamEtat;
  if Assigned(BParamEtat2)  then BParamEtat3.OnClick  := ParamEtat;

  if assigned(FETAT)        then FETAT.OnChange       := OnChangeEtat;
  if assigned(FETAT1)       then FETAT1.OnChange      := OnChangeEtat;
  if assigned(FETAT2)       then FETAT2.OnChange      := OnChangeEtat;
  if assigned(FETAT3)       then FETAT3.OnChange      := OnChangeEtat;

end;

procedure TOF_BTEDTBONPAIEMENT.VisibleScreenObject;
begin

  if fetat.Items.Count > 1 then
  begin
    Fetat.Visible := true;
    Tetat.Visible := true;
  end;

  GRB_FRAIS.visible := ChkEditjust.checked

end;

procedure TOF_BTEDTBONPAIEMENT.InitScreenObject;
begin

  ChkApercu.Checked   :=True;
  ChkReduire.Checked  :=False;
  ChkCouleur.Checked  :=False;
  ChkEditJust.Checked :=False;

  FEtat.Items.Clear;
  FEtat.Values.Clear;

  FEtat1.Items.Clear;
  FEtat1.Values.Clear;

  FEtat2.Items.Clear;
  FEtat2.Values.Clear;

  FEtat3.Items.Clear;
  FEtat3.Values.Clear;

end;

Procedure TOF_BTEDTBONPAIEMENT.ControleChamp(Champ : String;Valeur : String);
Begin

  Fgestion := TmgIntervenant;

  if champ = 'COTRAITANCE' then
  begin
    SetControlText('XX_WHERE', 'AND AFF_MANDATAIRE IN ("X", "-")');
    Fgestion := TmgCotraitance;
  end
  else if champ = 'SOUSTRAITANCE'then
  begin
    SetControlText('XX_WHERE', 'AND SSTRAITE > 0');
    Fgestion := TmgSousTraitance;
  end
  else if champ = 'INTERVENANT'then
  begin
    SetControlText('XX_WHERE', 'AND AFF_MANDATAIRE IN ("X", "-") OR SSTRAITE > 0');
    Fgestion := TmgIntervenant;
  end
  //
  else If Champ = 'STATUT' then
  Begin
    if Valeur = 'APP' then
    Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlText('TGP_AFFAIRE1', 'Appel');
    end
    Else if valeur = 'INT' then
    Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="I"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlText('TGP_AFFAIRE1', 'Contrat');
    end
    Else if valeur = 'AFF' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('GP_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) in ("A", "")');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlText('TGP_AFFAIRE1', 'Chantier');
    end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFFAIRE0')) then SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("W","A")');
    end
    Else if valeur = 'PRO' then
    Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("P")');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlText('TGP_AFFAIRE1', 'Appel d''offre');
    end
    else
    Begin
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlText('TGP_AFFAIRE1', 'Affaire');
    end
  end;

end;

procedure TOF_BTEDTBONPAIEMENT.ChargeEtatEdition;
begin

  ChargeEtatEdtEclatement;
  //
  ChargeEtatEdtDirecteCoTraite;
  //
  ChargeEtatEdtSsTraitant;
  //
  ChargeEtatEdtFrais;

end;

procedure TOF_BTEDTBONPAIEMENT.  ChargeEtatEdtEclatement;
begin

  TheType   := 'E';
  TheNature := 'BSO';
  TheModele := '';
  TheTitre  := 'Edition Lettre d''Eclatement';

  OptionEdition := TOptionEdition.Create(TheType, TheNature, TheModele, TheTitre, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, TPageControl(Ecran.FindComponent('Pages')), fEtat);

  OptionEdition.first := True;
  OptionEdition.ChargeListeEtat(fEtat,Idef);

  if fetat.itemindex  >= 0 then TheModele := FETAT.values[fetat.itemindex];

end;

procedure TOF_BTEDTBONPAIEMENT.ChargeEtatEdtDirecteCoTraite;
begin
  //
  TheType1   := 'E';
  TheNature1 := 'BSC';
  TheModele1 := '';
  TheTitre1  := 'Attestation Paiement Direct Cotraitant';

  OptionEdition1 := TOptionEdition.Create(TheType1, TheNature1, TheModele1, TheTitre1, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, TPageControl(Ecran.FindComponent('Pages')), fEtat1);
  OptionEdition1.first := True;
  OptionEdition1.ChargeListeEtat(fEtat1,Idef);

  if fetat1.itemindex  >= 0 then TheModele1 := FETAT1.values[fetat1.itemindex];

end;

procedure TOF_BTEDTBONPAIEMENT.ChargeEtatEdtSsTraitant;
begin
  //
  TheType2   := 'E';
  TheNature2 := 'BST';
  TheModele2 := '';
  TheTitre2  := 'Attestation Paiement Direct Sous-traitant';

  OptionEdition2 := TOptionEdition.Create(TheType2, TheNature2, TheModele2, TheTitre2, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, TPageControl(Ecran.FindComponent('Pages')), fEtat2);
  OptionEdition2.first := True;
  OptionEdition2.ChargeListeEtat(fEtat2,Idef);

  if fetat2.itemindex  >= 0 then TheModele2 := FETAT2.values[fetat2.itemindex];

end;

procedure TOF_BTEDTBONPAIEMENT.ChargeEtatEdtFrais;
begin
  //
  TheType3   := 'E';
  TheNature3 := 'BSR';
  TheModele3 := '';
  TheTitre3  := 'Récapitulatif des frais';

  OptionEdition3 := TOptionEdition.Create(TheType3, TheNature3, TheModele3, TheTitre3, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, TPageControl(Ecran.FindComponent('Pages')), fEtat3);
  OptionEdition3.first := True;
  OptionEdition3.ChargeListeEtat(fEtat3,Idef);

  if fetat3.itemindex  >= 0 then TheModele3 := FETAT3.values[fetat3.itemindex];

end;

procedure TOF_BTEDTBONPAIEMENT.CreateTobs;
Begin

  TobPIECE      := TOB.Create('PIECE',nil, -1);
  TobAFFAIRE    := TOB.Create('LES AFFAIRES',nil, -1);
  TobPiecTrait  := TOB.Create('PIECETRAIT',nil, -1);
  //
  TobSousTraite := TOB.Create('PIECEINTERV', nil, -1);
  //
  TobTiersFact  := TOB.Create('TIERSFACT', nil, -1);
  TobMandataire := TOB.Create('MANDATAIRE', nil, -1);
  TobSociete    := TOB.Create('SOCIETE', nil, -1);
  //
  TOBPieceRG := TOB.Create ('LES RETENUES',nil,-1);
  TOBBasesRg := TOB.Create ('LES BASESRG',nil,-1);
  //
  TobCalcul := TOB.Create('CALCUL', nil, -1);

  //chargement de la tob globale d'édition
  TobEditions := TOB.Create('LES EDITIONS', nil, -1);

  TOBEC := TOB.create ('ETATS COTRAITANTS',nil,-1);
  TOBES := TOB.create ('ETATS SOUSTRAITANTS',nil,-1);
  //
  TOBRecap := TOB.create ('LES RECAPS',nil,-1);
end;

procedure TOF_BTEDTBONPAIEMENT.ReleaseTob;
Begin
  //
  FreeAndNil(TobPIECE);
  FreeAndNil(TobAffaire);
  FreeAndNil(TobPiecTrait);
  FreeAndNil(TobSousTraite);
  //
  FreeAndNil(TobTiersFact);
  FreeAndNil(TobMandataire);
  FreeAndNil(TobSociete);
  //
  FreeAndnil(TobPieceRG);
  FreeAndNil(TOBBasesRg);
  //
  FreeAndNil(TobCalcul);
  //
  FreeAndNil(TobEditions);

  FreeAndNil(TOBEC);
  FreeAndNil(TOBES);
  FreeAndNil(TOBRecap);

end;

//appel de l'écran de modificatuion/création d'état
procedure TOF_BTEDTBONPAIEMENT.ParamEtat(Sender : TOBJect);
begin

  if TToolBarButton97(sender).name = 'BPARAMETAT' then
    OptionEdition.Appel_Generateur
  else if TToolBarButton97(sender).name = 'BPARAMETAT1' then
    OptionEdition1.Appel_Generateur
  else if TToolBarButton97(sender).name = 'BPARAMETAT2' then
    OptionEdition2.Appel_Generateur
  else if TToolBarButton97(sender).name = 'BPARAMETAT3' then
    OptionEdition3.Appel_Generateur;

end;

procedure TOF_BTEDTBONPAIEMENT.TraiteLesLignes(Sender : TOBJect);
var iInd        : Integer;
begin

  CreateTobs;

  Ok_Multiple := False;
  Ok_LastEdit := False;

  if not FListe.AllSelected then
  begin
    if FListe.NbSelected = 0 then
      PGIInfo('Selectionner au moins une ligne SVP !', 'Edition des bons de payement')
    else
    begin
      for iInd := 0 to FListe.NbSelected -1 do
      begin
        FListe.GotoLeBookMark(iInd);
        cledoc.NaturePiece  := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        cledoc.NumeroPiece  := Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        cledoc.Souche       := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        cledoc.Indice       := Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        //Chargement des Tob principales nécessaires pour l'édition
        ChargelesTobsPrincipales(Cledoc);
        //
        ConstitueTobCalcul;
        //
        ConstitueEditions;
      end;
      EditeTout;
    end;
  end;

  FListe.ClearSelected;

  TToolbarButton97(ecran.FindComponent('BOuvrir')).Down := False;
  FListe.AllSelected := False;

  ReleaseTob;

end;

Procedure TOF_BTEDTBONPAIEMENT.ChargelesTobsPrincipales(Cledoc : R_Cledoc);
Var StSQL : string;
    QQ    : TQuery;
begin

  //chargement de la pièce...
  StSQL := 'SELECT * FROM PIECE WHERE ' + WherePiece(Cledoc, ttdPiece,True);
  QQ:=OpenSQL(StSQL, True, 1, '', True);
  if not QQ.Eof then
  begin
    TobPiece.SelectDB('', QQ);
    //
    ChargementTiersFacture(TobPIECE.GetString('GP_TIERSFACTURE'));
    //
    //chargement des infos société
    ChargementSociete;
    //
    ChargementMandataire(TobPIECE);
    //
    //chargement informations cotraitance au niveau de la piece
    StSQL := 'Select *,'+
             '(SELECT BPI_TYPEPAIE FROM PIECEINTERV WHERE '+
             'BPI_NATUREPIECEG=BPE_NATUREPIECEG AND BPI_SOUCHE=BPE_SOUCHE AND '+
             'BPI_NUMERO=BPE_NUMERo AND BPI_INDICEG=BPE_INDICEG AND BPI_TIERSFOU=BPE_FOURNISSEUR) AS REGSSTRAIT FROM PIECETRAIT WHERE ' +
             WherePiece(cledoc, ttdPieceTrait, True);
    TobPiecTrait.LoadDetailFromSQL(StSQl);
  end;

  ferme(QQ);

  LoadLesRetenues(TOBPiece, TOBPieceRG, TOBBasesRG,taModif);

end;

procedure TOF_BTEDTBONPAIEMENT.ConstitueTobCalcul;
begin

  TobCalcul.Dupliquer(TobPiecTrait, True, true);
  TobCalcul.AddChampSupValeur('TODAY', Now);

  CreationLigneMandataire;

  ChargementligneFrais;

end;

Procedure TOF_BTEDTBONPAIEMENT.AddlesChampsSupMandataire (TOBmandataire : TOB);
begin

  TobMandataire.AddChampSupValeur('MANDATAIRE', '');
  TobMandataire.AddChampSupValeur('REGROUPEMENT', '');
  TobMandataire.AddChampSupValeur('CODEBQ', '');
  TobMandataire.AddChampSupValeur('BQMANDATAIRE', '');
  TobMandataire.AddChampSupValeur('NOMMANDAT', '');
  TobMandataire.AddChampSupValeur('ADRESSE1', '');
  TobMandataire.AddChampSupValeur('ADRESSE2', '');
  TobMandataire.AddChampSupValeur('ADRESSE3', '');
  TobMandataire.AddChampSupValeur('CODEPOSTAL', '');
  TobMandataire.AddChampSupValeur('VILLE', '');
  TobMandataire.AddChampSupValeur('TELEPHONE', '');
  TobMandataire.AddChampSupValeur('FAX', '');
  TobMandataire.AddChampSupValeur('EMAIL', '');
  TobMandataire.AddChampSupValeur('STATUTMANDAT', '');
  TobMandataire.AddChampSupValeur('SIRETMANDAT', '');
  TobMandataire.AddChampSupValeur('APEMANDAT', '');
  TobMandataire.AddChampSupValeur('TYPEPAIE', '');
  TobMandataire.AddChampSupValeur('CPT_REGROUPEMENT', '');
  TobMandataire.AddChampSupValeur('LIB_REGROUPEMENT', '');
  TobMandataire.AddChampSupValeur('ADR1_REGROUPEMENT','');
  TobMandataire.AddChampSupValeur('ADR2_REGROUPEMENT','');
  TobMandataire.AddChampSupValeur('ADR3_REGROUPEMENT','');
  TobMandataire.AddChampSupValeur('CP_REGROUPEMENT',  '');
  TobMandataire.AddChampSupValeur('DOM_REGROUPEMENT', '');
  TobMandataire.AddChampSupValeur('VIL_REGROUPEMENT', '');
  TobMandataire.AddChampSupValeur('DESBANQUE', '');
  TobMandataire.AddChampSupValeur('ETABBQ', '');
  TobMandataire.AddChampSupValeur('GUICHET','');
  TobMandataire.AddChampSupValeur('COMPTE_COTRAITE', '');
  TobMandataire.AddChampSupValeur('CLERIB', '');
  TobMandataire.AddChampSupValeur('DOMICILIATION', '');
  TobMandataire.AddChampSupValeur('VILLEBQ', '');
  TobMandataire.AddChampSupValeur('IBAN', '');

end;

Procedure TOF_BTEDTBONPAIEMENT.ChargementMandataire(TOBpiece : TOB);
var StSQL : string;
    QQ    : TQuery;
    TOBBQE: TOB;
    CodeAffaire,Etablissement : string;
begin
  CodeAffaire := TOBPiece.GetString('GP_AFFAIRE');
  AddlesChampsSupMandataire (TOBmandataire);
  //lecture de la fiche affaire pour chargement mandataire
  if CodeAffaire <> '' then
  begin
    StSql := 'Select AFF_BQMANDATAIRE, AFF_CODEBQ, AFF_TIERS, AFF_ETABLISSEMENT, AFF_REFEXTERNE, AFF_TYPEPAIE FROM AFFAIRE WHERE AFF_AFFAIRE="' + CodeAffaire +'"';
    QQ := OpenSQL(StSql, True);
    if not QQ.eof then
    begin
      //chargement des info mandataires associées à l'affaire...
      TobMandataire.SetString('MANDATAIRE', QQ.FindField('AFF_ETABLISSEMENT').AsString);

      TobMandataire.SetString('REGROUPEMENT', QQ.FindField('AFF_REFEXTERNE').AsString);
      TobMandataire.SetString('CODEBQ', QQ.FindField('AFF_CODEBQ').AsString);
      TobMandataire.SetString('BQMANDATAIRE', QQ.FindField('AFF_BQMANDATAIRE').AsString);
      //
      TobMandataire.SetString('TYPEPAIE', QQ.FindField('AFF_TYPEPAIE').AsString);
      Etablissement := QQ.FindField('AFF_ETABLISSEMENT').AsString;
    end;
  	Ferme(QQ);
  end else
  begin
  	Etablissement := TOBPiece.GetString('GP_ETABLISSEMENT');
  end;
  //lecture de l'établissement
  StSql := 'SELECT * FROM ETABLISS WHERE ET_ETABLISSEMENT="' + Etablissement +'"';
  QQ := OpenSQL(StSql, True);
  if not QQ.eof then
  begin
    //chargement des infos société
    TobMandataire.PutValue('NOMMANDAT',   QQ.FindField('ET_LIBELLE').AsString);
    TobMandataire.PutValue('ADRESSE1',    QQ.FindField('ET_ADRESSE1').AsString);
    TobMandataire.PutValue('ADRESSE2',    QQ.FindField('ET_ADRESSE2').AsString);
    TobMandataire.PutValue('ADRESSE3',    QQ.FindField('ET_ADRESSE3').AsString);
    TobMandataire.PutValue('CODEPOSTAL',  QQ.FindField('ET_CODEPOSTAL').AsString);
    TobMandataire.PutValue('VILLE',       QQ.FindField('ET_VILLE').AsString);
    TobMandataire.PutValue('TELEPHONE',   QQ.FindField('ET_TELEPHONE').AsString);
    TobMandataire.PutValue('FAX',         QQ.FindField('ET_FAX').AsString);
    TobMandataire.PutValue('EMAIL',       QQ.FindField('ET_EMAIL').AsString);
    TobMandataire.PutValue('STATUTMANDAT',QQ.FindField('ET_JURIDIQUE').AsString);
    TobMandataire.PutValue('SIRETMANDAT', QQ.FindField('ET_SIRET').AsString);
    TobMandataire.PutValue('APEMANDAT',   QQ.FindField('ET_APE').AsString);
  end else
  begin
    TobMandataire.PutValue('MANDATAIRE',  GetParamSocSecur('SO_SOCIETE', ''));
    TobMandataire.PutValue('NOMMANDAT',   GetParamSocSecur('SO_LIBELLE', ''));
  end;

  TOBBQE := Tob.Create('BANQUECP', nil, -1);
  if TobMandataire.GetString('CODEBQ') <> '' then
  begin
    //chargement du numero de compte regroupement...
    LectBanque(TOBBQE, TobMandataire.GetString('CODEBQ'));
    TobMandataire.SetString('CPT_REGROUPEMENT', TOBBQE.GetString('BQ_NUMEROCOMPTE'));
    TobMandataire.SetString('LIB_REGROUPEMENT', TOBBQE.GetString('BQ_LIBELLE'));
    TobMandataire.SetString('ADR1_REGROUPEMENT',TOBBQE.GetString('BQ_ADRESSE1'));
    TobMandataire.SetString('ADR2_REGROUPEMENT',TOBBQE.GetString('BQ_ADRESSE2'));
    TobMandataire.SetString('ADR3_REGROUPEMENT',TOBBQE.GetString('BQ_ADRESSE3'));
    TobMandataire.SetString('CP_REGROUPEMENT',  TOBBQE.GetString('BQ_CODEPOSTAL'));
    TobMandataire.SetString('DOM_REGROUPEMENT', TOBBQE.GetString('BQ_DOMICILIATION'));
    TobMandataire.SetString('VIL_REGROUPEMENT', TOBBQE.GetString('BQ_VILLE'));
  end;
  //chargement du compte bancaire du mandataire...
  if TobMandataire.GetString('BQMANDATAIRE') <> '' then
  begin
    LectBanque(TOBBQE, TobMandataire.GetString('BQMANDATAIRE'));

    TobMandataire.SetString('DESBANQUE', TOBBQE.GetString('BQ_LIBELLE'));
    TobMandataire.SetString('ETABBQ', TOBBQE.GetString('BQ_ETABBQ'));
    TobMandataire.SetString('GUICHET',TOBBQE.GetString('BQ_GUICHET'));
    TobMandataire.SetString('COMPTE_COTRAITE', TOBBQE.GetString('BQ_NUMEROCOMPTE'));
    TobMandataire.SetString('CLERIB', TOBBQE.GetString('BQ_CLERIB'));
    TobMandataire.SetString('DOMICILIATION', TOBBQE.GetString('BQ_DOMICILIATION'));
    TobMandataire.SetString('VILLEBQ', TOBBQE.GetString('BQ_VILLE'));
    TobMandataire.SetString('IBAN', TOBBQE.GetString('BQ_CODEIBAN'));
  end;
  FreeAndNil(TOBBQE);

end;

Procedure TOF_BTEDTBONPAIEMENT.ChargementSociete;
begin

  TobSociete.AddChampSupValeur('SOCIETE',     GetParamSocSecur('SO_SOCIETE', ''));
  TobSociete.AddChampSupValeur('RAISONSOC',   GetParamSocSecur('SO_LIBELLE', ''));
  TobSociete.AddChampSupValeur('ADRSOC1',     GetParamSocSecur('SO_ADRESSE1', ''));
  TobSociete.AddChampSupValeur('ADRSOC2',     GetParamSocSecur('SO_ADRESSE2',''));
  TobSociete.AddChampSupValeur('ADRSOC3',     GetParamSocSecur('SO_ADRESSE3',''));
  TobSociete.AddChampSupValeur('CPSOC',       GetParamSocSecur('SO_CODEPOSTAL',''));
  TobSociete.AddChampSupValeur('VILLESOC',    GetParamSocSecur('SO_VILLE', ''));
  TobSociete.AddChampSupValeur('TELSOC',      GetParamSocSecur('SO_TELEPHONE', ''));
  TobSociete.AddChampSupValeur('FAXSOC',      GetParamSocSecur('SO_FAX', ''));
  TobSociete.AddChampSupValeur('EMAILSOC',    GetParamSocSecur('SO_EMAIL', ''));
  TobSociete.AddChampSupValeur('JURIDIQUE',   GetParamSocSecur('SO_NATUREJURIDIQUE', ''));
  TobSociete.AddChampSupValeur('CAPITAL',     GetParamSocSecur('SO_CAPITAL', ''));
  TobSociete.AddChampSupValeur('SIRET',       GetParamSocSecur('SO_SIRET', ''));
  TobSociete.AddChampSupValeur('APE',         GetParamSocSecur('SO_APE', ''));
  TobSociete.AddChampSupValeur('RC',          GetParamSocSecur('SO_RC', ''));
  TobSociete.AddChampSupValeur('RVA',         GetParamSocSecur('SO_RVA', ''));

end;


Procedure TOF_BTEDTBONPAIEMENT.ChargementTiersFacture(TiersFacture : string);
var StSQL         : string;
    QQ            : TQuery;
begin

  //Recherche adresse tiers de la pièce facturé...
  StSql := 'Select * FROM PIECEADRESSE WHERE GPA_TYPEPIECEADR="001" AND ' + WherePiece(CleDoc, ttdPieceAdr, true);
  QQ := OpenSQL(StSql, True);
  if QQ.eof then
    PGIInfo('erreur de chargement de l''adresse du client Facturé :' + TiersFacture + ' du document !')
  else
  begin
    TobTiersFact.AddChampSupValeur('TIERSFACTURE', TiersFacture);
    TobTiersFact.AddChampSupValeur('NOMTIERS',    QQ.FindField('GPA_LIBELLE').AsString);
    TobTiersFact.AddChampSupValeur('ADRTIERS1',   QQ.FindField('GPA_ADRESSE1').AsString);
    TobTiersFact.AddChampSupValeur('ADRTIERS2',   QQ.FindField('GPA_ADRESSE2').AsString);
    TobTiersFact.AddChampSupValeur('ADRTIERS3',   QQ.FindField('GPA_ADRESSE3').AsString);
    TobTiersFact.AddChampSupValeur('CPTIERS',     QQ.FindField('GPA_CODEPOSTAL').AsString);
    TobTiersFact.AddChampSupValeur('VILLETIERS',  QQ.FindField('GPA_VILLE').AsString);
  end;
  Ferme(QQ);

end;

procedure TOF_BTEDTBONPAIEMENT.CreationLigneMandataire;
var TOBXX       : TOB;
begin

  //Création de la ligne mandataire
  TOBXX := TobCalcul.FindFirst(['BPE_FOURNISSEUR'], [''], false);

  if TOBXX <> nil then FreeAndNil(TOBXX);

  TOBXX := TOB.create('', TobCalcul, 0);

  TobXX.AddChampSupValeur('NATUREPIECEG', TobPIECE.GetString('GP_NATUREPIECEG'));
  TobXX.AddChampSupValeur('SOUCHE',  TobPIECE.GetString('GP_SOUCHE'));
  TobXX.AddChampSupValeur('INDICEG', TobPIECE.GetString('GP_INDICEG'));
  TobXX.AddChampSupValeur('AFFAIRE', TobPIECE.GetString('GP_AFFAIRE'));
  TobXX.AddChampSupValeur('BPE_FOURNISSEUR', '');
  TobXX.AddChampSupValeur('MANDATAIRE', TobMandataire.GetString('MANDATAIRE'));
  TobXX.AddChampSupValeur('NOMMANDAT', TobMandataire.GetString('NOMMANDAT'));

  TobXX.AddChampSupValeur('ETABBQ', TobMandataire.GetString('ETABBQ'));
  TobXX.AddChampSupValeur('GUICHET', TobMandataire.GetString('GUICHET'));
  TobXX.AddChampSupValeur('COMPTE_COTRAITE', TobMandataire.GetString('COMPTE_COTRAITE'));
  TobXX.AddChampSupValeur('CLERIB', TobMandataire.GetString('CLERIB'));
  TobXX.AddChampSupValeur('DOMICILIATION', TobMandataire.GetString('DOMICILIATION'));
  TobXX.AddChampSupValeur('VILLEBQ', TobMandataire.GetString('VILLEBQ'));

  TobXX.AddChampSupValeur('TYPEPAIE', '');

  TobXX.AddChampSupValeur('TOTALFRAIS', 0);
  TobXX.AddChampSupValeur('MTDEMANDE', 0);
  TobXX.AddChampSupValeur('MONTANTREGLABLE', 0);
  TobXX.AddChampSupValeur('MONTANTREGLE', 0);
  TobXX.AddChampSupValeur('TOTALRG', 0);   

end;

Procedure TOF_BTEDTBONPAIEMENT.Chargementlignefrais;
var iInd      : integer;
    jInd      : Integer;
    NumRib    : integer;
    //
    Prefixe   : string;
    TypeInter : string;
    StSql     : string;
    CodeFrs   : String;
    CodeAuxi  : String;
    Affaire   : String;
    NumCaution: String;
    //
    QQ        : TQuery;
    //
    CoefFrais : Double;
    TotDemande: Double;
    TotFrais  : Double;
    TotRegle  : Double;
    LigFrais  : Double;
    LigRegle  : Double;
    MtDemande : Double;
    MtFrais   : Double;
    LigRG     : Double;
    TotRG     : Double;
    Xp        : Double;
    //
    TOBBQE    : TOB;
    TOBRG     : TOB;
begin

  Affaire     := TobPIECE.GetString('GP_AFFAIRE');
  //
  TotDemande  := TobPIECE.GetDouble('GP_TOTALTTCDEV');
  TotFrais    := 0;
  TotRegle    := 0;
  TotRG       := 0;

  //mise en place de la Tob de calcul des montants réglés et à régler par cotraitant et mandataires
  for iInd := 0 to TobCalcul.Detail.Count -1 do
  begin
    CodeFrs   := TobCalcul.detail[iInd].GetString('BPE_FOURNISSEUR');
    CodeAuxi := TiersAuxiliaire (CodeFrs,false,'FOU');
    TypeInter := TobCalcul.detail[iInd].Getstring('BPE_TYPEINTERV');
    //
    //lecture des frais de gestion
    if CodeFrs  <> '' then
    begin
      if (TypeInter = 'X00') Or (TypeInter = 'X01') then
        MtDemande := TobCalcul.detail[iInd].GetDouble('BPE_TOTALTTCDEV')
      else
        MtDemande := TobCalcul.detail[iInd].GetDouble('BPE_MONTANTPA');

      LigFrais := 0;
      LigRG    := 0;

      //Création des zone pour valeurs cotraitant
      TobCalcul.Detail[iInd].AddChampSupValeur('FOURNISSEUR', CodeFrs);
      TobCalcul.Detail[iInd].AddChampSupValeur('NOMFRS',        '');
      TobCalcul.Detail[iInd].AddChampSupValeur('ADRESSE1',      '');
      TobCalcul.Detail[iInd].AddChampSupValeur('ADRESSE2',      '');
      TobCalcul.Detail[iInd].AddChampSupValeur('ADRESSE3',      '');
      TobCalcul.Detail[iInd].AddChampSupValeur('CODEPOSTAL',    '');
      TobCalcul.Detail[iInd].AddChampSupValeur('VILLE',         '');
      TobCalcul.Detail[iInd].AddChampSupValeur('DESBANQUE',     '');
      TobCalcul.Detail[iInd].AddChampSupValeur('ETABBQ',        '');
      TobCalcul.Detail[iInd].AddChampSupValeur('GUICHET',       '');
      TobCalcul.Detail[iInd].AddChampSupValeur('COMPTE_COTRAITE','');
      TobCalcul.Detail[iInd].AddChampSupValeur('CLERIB',        '');
      TobCalcul.Detail[iInd].AddChampSupValeur('DOMICILIATION', '');
      TobCalcul.Detail[iInd].AddChampSupValeur('VILLEBQ',       '');
      TobCalcul.Detail[iInd].AddChampSupValeur('IBAN',          '');
      TobCalcul.Detail[iInd].AddChampSupValeur('MTDEMANDE',       FormatFloat('#0.00', 0));
      TobCalcul.Detail[iInd].AddChampSupValeur('TOTALFRAIS',      FormatFloat('#0.00', 0));
      TobCalcul.Detail[iInd].AddChampSupValeur('MONTANTREGLABLE', FormatFloat('#0.00', 0));
      TobCalcul.Detail[iInd].AddChampSupValeur('MONTANTREGLE',    FormatFloat('#0.00', 0));
      TobCalcul.Detail[iInd].AddChampSupValeur('TOTALRG',         FormatFloat('#0.00', 0));
      TobCalcul.Detail[iInd].AddChampSupValeur('TYPEPAIE',      '');
      //
      //recherche informations fournisseurs (nom)
      StSQl := 'SELECT T_LIBELLE, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, T_CODEPOSTAL, T_VILLE From TIERS WHERE T_TIERS="' + CodeFrs + '"';
      QQ := OpenSQL(StSql, True);
      if not QQ.Eof then
      begin
        TobCalcul.Detail[iInd].PutValue('NOMFRS',     QQ.Findfield('T_LIBELLE').asString);
        TobCalcul.Detail[iInd].PutValue('ADRESSE1',   QQ.FindField('T_ADRESSE1').AsString);
        TobCalcul.Detail[iInd].PutValue('ADRESSE2',   QQ.FindField('T_ADRESSE2').AsString);
        TobCalcul.Detail[iInd].PutValue('ADRESSE3',   QQ.FindField('T_ADRESSE3').AsString);
        TobCalcul.Detail[iInd].PutValue('CODEPOSTAL', QQ.FindField('T_CODEPOSTAL').AsString);
        TobCalcul.Detail[iInd].PutValue('VILLE',      QQ.FindField('T_VILLE').AsString);
      end
      Else
      begin
        TobCalcul.Detail[iInd].PutValue('NOMFRS', 'Intervenant inconnu');
      end;
      Ferme(QQ);

      if TypeInter ='Y00' then
      begin
        Prefixe := 'BPI_';
        StSql := 'SELECT BPI_TYPEPAIE, BPI_NUMERORIB FROM PIECEINTERV WHERE AND BPI_TIERSFOU ="' + CodeFrs + '" AND ' + WherePiece(cledoc, ttdPieceInterv, True);;
      end
      else
      begin
        //Recherche des ligne d'intervention sur affaire pour récupération renseignement bancaire....
        Prefixe := 'BAI_';
        StSql := 'SELECT BAI_NUMERORIB FROM AFFAIREINTERV WHERE BAI_AFFAIRE="' + Affaire + '" AND BAI_TIERSFOU ="' + CodeFrs + '"';
      end;

      QQ := OpenSQL(StSql, True);
      if not QQ.eof then
      begin
        NumRib := QQ.findfield(prefixe + 'NUMERORIB').AsInteger;
        if TypeInter ='Y00' then
          TobCalcul.Detail[iInd].PutValue('TYPEPAIE', QQ.findfield('BPI_TYPEPAIE').AsString);
        ferme (QQ);
      end
      else NumRib := 0;
      //
      //recherche des informations bancaires du cotraitant
      TOBBQE := Tob.Create('RIB', nil, -1);
      LectRIB(TOBBQE, CodeAuxi, NumRIB);
      TobCalcul.Detail[iInd].PutValue('DESBANQUE', '');
      TobCalcul.Detail[iInd].PutValue('ETABBQ',          TOBBQE.GetString('R_ETABBQ'));
      TobCalcul.Detail[iInd].PutValue('GUICHET',         TOBBQE.GetString('R_GUICHET'));
      TobCalcul.Detail[iInd].PutValue('COMPTE_COTRAITE', TOBBQE.GetString('R_NUMEROCOMPTE'));
      TobCalcul.Detail[iInd].PutValue('CLERIB',          TOBBQE.GetString('R_CLERIB'));
      TobCalcul.Detail[iInd].PutValue('DOMICILIATION',   TOBBQE.GetString('R_DOMICILIATION'));
      TobCalcul.Detail[iInd].PutValue('VILLEBQ',         TOBBQE.GetString('R_VILLE'));
      TobCalcul.Detail[iInd].PutValue('IBAN',            TOBBQE.GetString('R_CODEIBAN'));
      //
      //Recherche des frais de gestion sur cotraitant...
      StSQl := 'Select  BAF_AFFAIRE, BAF_TIERS, BAF_CODEPORT, BAF_TYPEPORT, ';
      StSQL := StSQL + 'GPO_LIBELLE, BAF_COEFF, BAF_PVHT, BAF_PVTTC';
      StSQL := StSQL + ' FROM AFFAIREFRSGEST LEFT JOIN PORT ON BAF_CODEPORT=GPO_CODEPORT ';
      StSQL := StSQL + 'WHERE BAF_AFFAIRE="' + Affaire + '"';
      StSql := StSql + '  AND BAF_TIERS="' + CodeFrs + '"';

      TobCalcul.Detail[Iind].LoadDetailFromSQL(StSQL);

      For jInd := 0 to TobCalcul.Detail[iInd].Detail.count -1 do
      begin
        MtFrais := TobCalcul.Detail[iInd].Detail[jInd].GetDouble('BAF_PVTTC');
        if MtFrais = 0 then MtFrais := TobCalcul.Detail[iInd].Detail[jInd].GetDouble('BAF_PVHT');
        CoefFrais := TobCalcul.Detail[iInd].Detail[jInd].GetDouble('BAF_COEFF');
        if CoefFrais <> 0 then MtFrais := (MtDemande * (CoefFrais/100));
        TobCalcul.Detail[iInd].Detail[jInd].AddChampSupValeur('MTFRAIS', FormatFloat('#0.00', MtFrais));
        LigFrais := LigFrais + MtFrais;
      end;

      //
      LigRegle := TobCalcul.detail[iInd].GetDouble('BPE_MONTANTREGL') - LigFrais;
      //

      //chargement des valeurs des montants de Retenue Garantie
      if TypeInter <>'Y00' then
      begin
        GetRG(TOBPieceRG, false, true, Xp, LigRG, NumCaution,CodeFrs,False);
        if LigRG <> 0 then
        begin
          TOBRG := Tob.create('RETENUE GARANTIE',TobCalcul.Detail[iInd], -1);
          TOBRG.AddChampSupValeur('BAF_AFFAIRE', Affaire);
          TOBRG.AddChampSupValeur('BAF_TIERS', CodeFrs);
          TOBRG.AddChampSupValeur('BAF_CODEPORT', '');
          TOBRG.AddChampSupValeur('BAF_TYPEPORT', 'MT');
          TOBRG.AddChampSupValeur('GPO_LIBELLE', 'Retenue de Garantie');
          TOBRG.AddChampSupValeur('BAF_COEFF', FormatFloat('#0.00', 0));
          TOBRG.AddChampSupValeur('BAF_PVHT',  FormatFloat('#0.00', 0));
          TOBRG.AddChampSupValeur('BAF_PVTTC', FormatFloat('#0.00', LigRG));
          LigRegle := LigRegle - LigRG;
        end;
      end;
      //

      TobCalcul.Detail[iInd].PutValue('MTDEMANDE',       FormatFloat('#0.00',MtDemande));
      TobCalcul.Detail[iInd].PutValue('TOTALFRAIS',      FormatFloat('#0.00',LigFrais));
      TobCalcul.Detail[iInd].PutValue('MONTANTREGLABLE', FormatFloat('#0.00',MTDemande - LigFrais));
      TobCalcul.Detail[iInd].PutValue('MONTANTREGLE',    FormatFloat('#0.00',LigRegle));
      TobCalcul.Detail[iInd].PutValue('TOTALRG',         FormatFloat('#0.00',LigRg));

      TotFrais    := TotFrais + LigFrais;
      TotRegle    := TotRegle + LigRegle;
      TotRG       := TotRG    + LigRG;
    end;
  end;

  //repositionnement sur le mandataire pour calcul du montant réglé...
  TotRegle := TotDemande - TotRegle;

  TobCalcul.Detail[0].PutValue('TOTALFRAIS',      FormatFloat('#0.00',TotFrais));
  TobCalcul.Detail[0].PutValue('MTDEMANDE',       FormatFloat('#0.00',TotDemande));
  TobCalcul.Detail[0].PutValue('MONTANTREGLABLE', FormatFloat('#0.00',TotRegle));
  TobCalcul.Detail[0].PutValue('MONTANTREGLE',    FormatFloat('#0.00',TotRegle));
  TobCalcul.Detail[0].PutValue('TOTALRG',         FormatFloat('#0.00',TotRg));

end;


procedure TOF_BTEDTBONPAIEMENT.ConstitueEditions;
var iInc      : Integer;
    SaveiInc  : Integer;

    TotalMan  : Double;

    TobEditC  : TOB;
    TobEditS  : TOB;
    TobCalc   : TOB;

    Ok_EnteteSST : Boolean;
    Ok_EnteteCOT : Boolean;

    Ok_Direct : Boolean;
Begin

  Ok_EnteteSST := False;
  Ok_EnteteCOT := False;

  Ok_Recap  := False;

  //
  //chargement de la tob d'édition par la tob de Calcul....
  TobEditC := TOB.Create('UNE EDITION', TOBEC, -1);
  TobEditS := Tob.create('UNE EDITION', TOBES, -1);

  TRY
    iInc    := 0;
    SaveiInc:= 0;
    TotalMan:= 0;

    Repeat
      TobCalc := TOBCalcul.detail[iInc];
      if TobCalc.Getstring('REGSSTRAIT')='' then
      begin
        if TobMandataire.GetString('TYPEPAIE') = '000' then
          Ok_Direct := False
        else
          Ok_Direct := True;
      end else
      begin
        Ok_Direct := (TOBCalc.GetValue('REGSSTRAIT')='001');
      end;
      if not Ok_Direct then      //Regroupement sur paiement
      begin
        if not Ok_EnteteCOT then
        begin
          ChargeEnteteEdition(TobEditC, TobMandataire.GetString('TYPEPAIE'));
          Ok_EnteteCOT := true;
        end;
        if TOBCalc.GetString('BPE_FOURNISSEUR') = '' then
        begin
          SaveiInc := iInc;
          TotalMan := TobCalcul.Detail[SaveiInc].GetDouble('MONTANTREGLE');
        end;
        EditionRegroupement(TobEditC, TobCalc, TotalMan);
      end
      else                       //Paiement direct
      begin
        //chargement de l'entête d'édition
        if TobCalc.GetString('BPE_TYPEINTERV') = 'Y00' then
        begin
          Ok_EdtSsTrait := True;
          if not Ok_EnteteSST then
          begin
            ChargeEnteteEdition(TobEditS, TobCalc.GetString('TYPEPAIE'));
            Ok_EnteteSST := True;
          end;
          EditionDirecte(TobEditS, TobCalc);
        end
        else
        begin
          if not Ok_EnteteCOT then
          begin
            ChargeEnteteEdition(TobEditC, TobMandataire.GetString('TYPEPAIE'));
            Ok_EnteteCOT := True;
          end;
          EditionDirecte(TobEditC, TobCalc);
        end;
      end;
      Inc(iInc) ;
    Until iInc > TobCalcul.detail.count-1;
  FINALLY
    PrepareEditions(TOBEC, TOBES);
  END;

end;


Procedure TOF_BTEDTBONPAIEMENT.PrepareEditions(TOBEC, TOBES : TOB);
Var TobEdtS : TOB;
    TobEdtST : TOB;
    TobEDtC : TOB;
    TobEdtR : TOB;
begin

  if TOBEC = nil then Exit;
  if TOBES = nil then Exit;

  //cotraitant ==> Regroupement
  TobEDtC := TOB.Create('ENTETE REG', TOBEditions, -1);
  AjouteLesParamsEdt (TobEDtC,OptionEdition);

  //sous-traitant ==> edition Directe
  TobEDtST := TOB.Create('ENTETE DIR ST', TOBEditions, -1);
  AjouteLesParamsEdt (TobEDtST,OptionEdition2);

  //Cotraitant ==> edition Directe
  TobEDtS := TOB.Create('ENTETE DIR', TOBEditions, -1);
  AjouteLesParamsEdt (TobEDtS,OptionEdition1);

  //Récapitulatif des frais
  TobEDtR := TOB.Create('ENTETE RECAP', TOBEditions, -1);
  AjouteLesParamsEdt (TobEDtR,OptionEdition3);

  //traitement des éditions Cotraitant (TOBEC)
  if TOBEC.Detail[0].detail.count > 1 then PrepareTOBEDit(TOBEc,TobedtC,TOBEdtS);

  if TOBES.Detail[0].detail.count > 0 then PrepareTOBEDit(TOBES,TobedtC,TOBEdtST);

  //Traitement des Editions recapitulatif de frais
  if ChkEditjust.Checked then
  begin
    GenereRecap;
    PrepareTOBRecap (TOBRecap,TOBEdtR);
  end;

end;

Procedure TOF_BTEDTBONPAIEMENT.EditionRegroupement(TobEdition, TobCalcul : TOB; TotalMan :Double);
Var TobLigEdt : Tob;
begin

  if not Assigned(TobEdition) then exit;

  if TOBCalcul.GetString('BPE_FOURNISSEUR') = '' then        //=> Gestion du mandataire
  begin
    //chargement des lignes mandataire dans la TobEdition
    TOBLigEdt := TOB.create('UNE LIGNE', TobEdition, -1);
    ChargeMandataireEdition(TobLigEdt, TOBCalcul);
  end
  else
  begin
    if TOBCalcul.GetString('BPE_TYPEINTERV') <> 'Y00' then //=> Gestion des cotraitants
      ChargeIntervenantEdition(TobEdition, TOBCalcul)
    else                                                   //=> Gestion des sous-traitants
    begin
      TotalMan := TotalMan + TobCalcul.GetDouble('MONTANTREGLE');
      TobEdition.detail[0].PutValue('MONTANTREGLE', TotalMan);
    end;
  end;

end;

Procedure TOF_BTEDTBONPAIEMENT.EditionDirecte(TobEdition, TobCalcul : TOB);
Var TobLigEdt : Tob;
begin

  if not Assigned(TobEdition) then exit;

  if TOBCalcul.GetString('BPE_FOURNISSEUR') = '' then
  begin
    TOBLigEdt := TOB.create('UNE LIGNE', TobEdition, -1);
    ChargeMandataireEdition(TobLigEdt, TOBCalcul);
  end
  else
  begin
    if TOBCalcul.GetString('BPE_TYPEINTERV') <> 'Y00' then  //=> Gestion des cotraitants
    begin
      ChargeIntervenantEdition(TobEdition, TOBCalcul);
    end
    else                                                    //=> Gestion des sous-traitants.
    begin
      if TOBCalcul.GetString('TYPEPAIE') = '001' then ChargeIntervenantEdition(TobEdition, TOBCalcul);
    end;
  end;

end;

Procedure TOF_BTEDTBONPAIEMENT.ChargeEnteteEdition(TobEdition : TOB; TypeInter : String);
Begin

  //chargement des informations du tiers facturé
  TobEdition.AddChampSupValeur('NOMTIERS',    TobTiersFact.GetString('NOMTIERS'));
  TobEdition.AddChampSupValeur('ADRTIERS1',   TobTiersFact.GetString('ADRTIERS1'));
  TobEdition.AddChampSupValeur('ADRTIERS2',   TobTiersFact.GetString('ADRTIERS2'));
  TobEdition.AddChampSupValeur('ADRTIERS3',   TobTiersFact.GetString('ADRTIERS3'));
  TobEdition.AddChampSupValeur('CPTIERS',     TobTiersFact.GetString('CPTIERS'));
  TobEdition.AddChampSupValeur('VILLETIERS',  TobTiersFact.GetString('VILLETIERS'));

  //chargement des informations de la société
  TobEdition.AddChampSupValeur('SOCIETE',     TobSociete.GetString('SOCIETE'));
  TobEdition.AddChampSupValeur('RAISONSOC',   TobSociete.GetString('RAISONSOC'));
  TobEdition.AddChampSupValeur('ADRSOC1',     TobSociete.GetString('ADRSOC1'));
  TobEdition.AddChampSupValeur('ADRSOC2',     TobSociete.GetString('ADRSOC2'));
  TobEdition.AddChampSupValeur('ADRSOC3',     TobSociete.GetString('ADRSOC3'));
  TobEdition.AddChampSupValeur('CPSOC',       TobSociete.GetString('CPSOC'));
  TobEdition.AddChampSupValeur('VILLESOC',    TobSociete.GetString('VILLESOC'));
  TobEdition.AddChampSupValeur('TELSOC',      TobSociete.GetString('TELSOC'));
  TobEdition.AddChampSupValeur('FAXSOC',      TobSociete.GetString('FAXSOC'));
  TobEdition.AddChampSupValeur('EMAILSOC',    TobSociete.GetString('EMAILSOC'));
  TobEdition.AddChampSupValeur('JURIDIQUE',   TobSociete.GetString('JURIDIQUE'));
  TobEdition.AddChampSupValeur('CAPITAL',     TobSociete.GetString('CAPITAL'));
  TobEdition.AddChampSupValeur('SIRET',       TobSociete.GetString('SIRET'));
  TobEdition.AddChampSupValeur('APESOC',      TobSociete.GetString('APE'));
  TobEdition.AddChampSupValeur('RC',          TobSociete.GetString('RC'));
  TobEdition.AddChampSupValeur('RVA',         TobSociete.GetString('RVA'));           

  //chargement des informations du mandataire/Etablissment
  TobEdition.AddChampSupValeur('MANDATAIRE',    TobMandataire.GetString('NOMMANDAT'));
  TobEdition.AddChampSupValeur('ADRMANDAT1',    TobMandataire.GetString('ADRESSE1'));
  TobEdition.AddChampSupValeur('ADRMANDAT2',    TobMandataire.GetString('ADRESSE2'));
  TobEdition.AddChampSupValeur('ADRMANDAT3',    TobMandataire.GetString('ADRESSE3'));
  TobEdition.AddChampSupValeur('CPMANDAT',      TobMandataire.GetString('CODEPOSTAL'));
  TobEdition.AddChampSupValeur('VILLEMANDAT',   TobMandataire.GetString('VILLE'));
  TobEdition.AddChampSupValeur('TELMANDAT',     TobMandataire.GetString('TELEPHONE'));
  TobEdition.AddChampSupValeur('FAXMANDAT',     TobMandataire.GetString('FAX'));
  TobEdition.AddChampSupValeur('EMAILMANDAT',   TobMandataire.GetString('EMAIL'));
  TobEdition.AddChampSupValeur('REGROUPEMENT',  TobMandataire.GetString('REGROUPEMENT'));
  TobEdition.AddChampSupValeur('CODEBANQUE',    TobMandataire.GetString('CODEBQ'));
  TobEdition.AddChampSupValeur('CPT_REGROUPEMENT',  TobMandataire.GetString('CPT_REGROUPEMENT'));
  TobEdition.AddChampSupValeur('LIB_REGROUPEMENT',  TobMandataire.GetString('LIB_REGROUPEMENT'));
  TobEdition.AddChampSupValeur('ADR1_REGROUPEMENT', TobMandataire.GetString('ADR1_REGROUPEMENT'));
  TobEdition.AddChampSupValeur('ADR2_REGROUPEMENT', TobMandataire.GetString('ADR2_REGROUPEMENT'));
  TobEdition.AddChampSupValeur('ADR3_REGROUPEMENT', TobMandataire.GetString('ADR3_REGROUPEMENT'));
  TobEdition.AddChampSupValeur('CP_REGROUPEMENT',   TobMandataire.GetString('CP_REGROUPEMENT'));

  TobEdition.AddChampSupValeur('DOM_REGROUPEMENT',  TobMandataire.GetString('DOM_REGROUPEMENT'));
  TobEdition.AddChampSupValeur('VIL_REGROUPEMENT',  TobMandataire.GetString('VIL_REGROUPEMENT'));

  TobEdition.AddChampSupValeur('STATUTMANDAT',  TobMandataire.GetString('STATUTMANDAT'));
  TobEdition.AddChampSupValeur('SIRETMANDAT',   TobMandataire.GetString('SIRETMANDAT'));
  TobEdition.AddChampSupValeur('APEMANDAT',     TobMandataire.GetString('APEMANDAT'));

  TobEdition.AddChampSupValeur('TYPEPAIE',      TypeInter);

  //Chargement des informations de la pièce facture
  TobEdition.AddChampSupValeur('NATUREPIECEG',  TobPiece.GetString('GP_NATUREPIECEG'));
  TobEdition.AddChampSupValeur('NUMERO',        TobPiece.GetString('GP_NUMERO'));
  TobEdition.AddChampSupValeur('DATEPIECE',     TobPiece.GetDateTime('GP_DATEPIECE'));
  TobEdition.AddChampSupValeur('TOTALTTCDEV',   TobPIECE.GetDouble('GP_TOTALTTCDEV'));
end;

Procedure TOF_BTEDTBONPAIEMENT.ChargeMandataireEdition(TobLigEdt, Tobcalc : TOB);
begin

  TOBLigEdt.AddChampSupValeur('TYPETIERS', 'M');
  TOBLigEdt.AddChampSupValeur('BAF_TIERS',        Tobcalc.GetString('BPE_FOURNISSEUR'));

  TOBLigEdt.AddChampSupValeur('NOMFRS',           TobMandataire.GetString('NOMMANDAT'));
  //
  TOBLigEdt.AddChampSupValeur('ADRESSE1',         TobMandataire.GetString('ADRESSE1'));
  TOBLigEdt.AddChampSupValeur('ADRESSE2',         TobMandataire.GetString('ADRESSE2'));
  TOBLigEdt.AddChampSupValeur('ADRESSE3',         TobMandataire.GetString('ADRESSE3'));
  TOBLigEdt.AddChampSupValeur('CODEPOSTAL',       TobMandataire.GetString('CODEPOSTAL'));
  TOBLigEdt.AddChampSupValeur('VILLE',            TobMandataire.GetString('VILLE'));
  //
  TOBLigEdt.AddChampSupValeur('NOMBQE',           TobMandataire.GetString('DESBANQUE'));
  TOBLigEdt.AddChampSupValeur('ETABBQ',           TobMandataire.GetString('ETABBQ'));
  TOBLigEdt.AddChampSupValeur('GUICHET',          TobMandataire.GetString('GUICHET'));
  TOBLigEdt.AddChampSupValeur('COMPTE_COTRAITE',  TobMandataire.GetString('COMPTE_COTRAITE'));
  TOBLigEdt.AddChampSupValeur('CLERIB',           TobMandataire.GetString('CLERIB'));
  TOBLigEdt.AddChampSupValeur('DOMICILIATION',    TobMandataire.GetString('DOMICILIATION'));
  TOBLigEdt.AddChampSupValeur('VILLEBQ',          TobMandataire.GetString('VILLEBQ'));
  TOBLigEdt.AddChampSupValeur('IBAN',             TobMandataire.GetString('IBAN'));

  TOBLigEdt.AddChampSupValeur('TYPEPAIE',         TobMandataire.GetString('TYPEPAIE'));
  TOBLigEdt.AddChampSupValeur('TOTALFRAIS',       Tobcalc.GetDouble('TOTALFRAIS'));
  TOBLigEdt.AddChampSupValeur('MTDEMANDE',        Tobcalc.GetDouble('MTDEMANDE'));
  TOBLigEdt.AddChampSupValeur('MONTANTREGLABLE',  Tobcalc.GetDouble('MONTANTREGLABLE'));
  TOBLigEdt.AddChampSupValeur('MONTANTREGLE',     Tobcalc.GetDouble('MONTANTREGLE'));
  TOBLigEdt.AddChampSupValeur('TOTALRG',          Tobcalc.GetDouble('TOTALRG'));

end;

Procedure TOF_BTEDTBONPAIEMENT.ChargeIntervenantEdition(TobMere, TobCalc : TOB);
Var TobLigEdt : Tob;
begin

  TOBLigEdt := TOB.create('UNE LIGNE', TobMere, -1);

  if TobCalc.GetString('BPE_TYPEINTERV') <> 'Y00' then
     TOBLigEdt.AddChampSupValeur('TYPETIERS', 'S')
  else
     TOBLigEdt.AddChampSupValeur('TYPETIERS', 'C');

  TOBLigEdt.AddChampSupValeur('BAF_TIERS',  TOBCalc.GetString('BPE_FOURNISSEUR'));

  TOBLigEdt.AddChampSupValeur('NOMFRS',     TOBCalc.GetString('NOMFRS'));
  //
  TOBLigEdt.AddChampSupValeur('ADRESSE1',   TobCalc.GetString('ADRESSE1'));
  TOBLigEdt.AddChampSupValeur('ADRESSE2',   TobCalc.GetString('ADRESSE2'));
  TOBLigEdt.AddChampSupValeur('ADRESSE3',   TobCalc.GetString('ADRESSE3'));
  TOBLigEdt.AddChampSupValeur('CODEPOSTAL', TobCalc.GetString('CODEPOSTAL'));
  TOBLigEdt.AddChampSupValeur('VILLE',      TobCalc.GetString('VILLE'));

  //chargement des informations du tiers facturé
  TOBLigEdt.AddChampSupValeur('NOMTIERS',   TobTiersFact.GetString('NOMTIERS'));
  TOBLigEdt.AddChampSupValeur('ADRTIERS1',  TobTiersFact.GetString('ADRTIERS1'));
  TOBLigEdt.AddChampSupValeur('ADRTIERS2',  TobTiersFact.GetString('ADRTIERS2'));
  TOBLigEdt.AddChampSupValeur('ADRTIERS3',  TobTiersFact.GetString('ADRTIERS3'));
  TOBLigEdt.AddChampSupValeur('CPTIERS',    TobTiersFact.GetString('CPTIERS'));
  TOBLigEdt.AddChampSupValeur('VILLETIERS', TobTiersFact.GetString('VILLETIERS'));
  //
  TOBLigEdt.AddChampSupValeur('NOMBQE',     TobCalc.GetString('DESBANQUE'));

  TOBLigEdt.AddChampSupValeur('ETABBQ',          TOBCalc.GetString('ETABBQ'));
  TOBLigEdt.AddChampSupValeur('GUICHET',         TOBCalc.GetString('GUICHET'));
  TOBLigEdt.AddChampSupValeur('COMPTE_COTRAITE', TOBCalc.GetString('COMPTE_COTRAITE'));
  TOBLigEdt.AddChampSupValeur('CLERIB',          TOBCalc.GetString('CLERIB'));
  TOBLigEdt.AddChampSupValeur('DOMICILIATION',   TOBCalc.GetString('DOMICILIATION'));
  TOBLigEdt.AddChampSupValeur('VILLEBQ',         TOBCalc.GetString('VILLEBQ'));
  TOBLigEdt.AddChampSupValeur('IBAN',            TOBCalc.GetString('IBAN'));

  TOBLigEdt.AddChampSupValeur('TOTALTTCDEV',     TobPIECE.GetDouble('GP_TOTALTTCDEV'));

  //Modif Type Paiement si sous-traitant
  TOBLigEdt.AddChampSupValeur('TYPEPAIE',        TOBCalc.GetString('TYPEPAIE'));

  TOBLigEdt.AddChampSupValeur('NUMERO',          TobPiece.GetString('GP_NUMERO'));

  TOBLigEdt.AddChampSupValeur('TOTALFRAIS',      TOBCalc.GetDouble('TOTALFRAIS'));
  TOBLigEdt.AddChampSupValeur('MTDEMANDE',       TOBCalc.GetDouble('MTDEMANDE'));
  TOBLigEdt.AddChampSupValeur('MONTANTREGLABLE', TOBCalc.GetDouble('MONTANTREGLABLE'));
  TOBLigEdt.AddChampSupValeur('MONTANTREGLE',    TOBCalc.GetDouble('MONTANTREGLE'));
  TOBLigEdt.AddChampSupValeur('TOTALRG',         TOBCalc.GetDouble('TOTALRG'));

end;

Procedure TOF_BTEDTBONPAIEMENT.ChargeRecapEdition(TobRecap, TobCalc : TOB);
var TobLigFrais : TOB;
begin

  //chargement des lignes de frais...
  TobLigFrais := TOB.Create('LES FRAIS', TobRecap, -1);

  TobLigFrais.AddChampSupValeur('BAF_CODEPORT', TOBCalc.GetString('BAF_CODEPORT'));
  TobLigFrais.AddChampSupValeur('LIBFRAIS',     TOBCalc.GetString('GPO_LIBELLE'));

  if TOBCalc.Getdouble('BAF_PVTTC') = 0 then
    TobLigFrais.AddChampSupValeur('MTFRAIS',    TOBCalc.GetDOUBLE('BAF_PVHT'))
  else
    TobLigFrais.AddChampSupValeur('MTFRAIS',    TOBCalc.GetDOUBLE('BAF_PVTTC'));

  TobLigFrais.AddChampSupValeur('PERCENT',      TOBCalc.GetDOUBLE('BAF_COEFF'));
  TobLigFrais.AddChampSupValeur('CALCFRAIS',    TOBCalc.GetDOUBLE('MTFRAIS'));

end;


Procedure TOF_BTEDTBONPAIEMENT.GenereRecap;
var iInc      : Integer;
    jInc      : Integer;
    NumLig    : Integer;
begin

  Ok_Recap := true;

  for iInc := 0 to Tobcalcul.Detail.Count -1 do
  begin
    NumLig := iInc-1;
    if TOBCalcul.detail[iInc].GetString('FOURNISSEUR') <> '' then
    begin
      ChargeIntervenantEdition(TobRecap, TOBCalcul.detail[iInc]); //chargement ligne récap
      For jInc := 0 To tobcalcul.Detail[iInc].detail.count -1 do
      begin
        ChargeRecapEdition(TobRecap.detail[NumLig], TOBCalcul.detail[iInc].detail[Jinc]); //chargement ligne de frais
      end;
    end;
  end;

end;

Procedure TOF_BTEDTBONPAIEMENT.OnChangeEtat(Sender : Tobject);
Begin

  if THValComboBOx(sender).name = 'FETAT' then
  begin
   OptionEdition.Modele := FETAT.values[fetat.itemindex];
   TheModele  := OptionEdition.Modele;
  end
  else if THValComboBOx(sender).name = 'FETAT1' then
  begin
   OptionEdition1.Modele := FETAT1.values[fetat1.itemindex];
   TheModele1 := OptionEdition1.Modele;
  end
  else if THValComboBOx(sender).name = 'FETAT2' then
  begin
   OptionEdition2.Modele := FETAT2.values[fetat2.itemindex];
   TheModele2 := OptionEdition2.Modele;
  end
  Else if THValComboBOx(sender).name = 'FETAT3' then
  begin
   OptionEdition3.Modele := FETAT3.values[fetat3.itemindex];
   TheModele3 := OptionEdition3.Modele;
  end

end;

procedure TOF_BTEDTBONPAIEMENT.OnClickApercu(Sender : Tobject);
begin
  OptionEdition.Apercu  := ChkApercu.checked;
  OptionEdition1.Apercu := ChkApercu.checked;
  OptionEdition2.Apercu := ChkApercu.checked;
  OptionEdition3.Apercu := ChkApercu.checked;
end;

procedure TOF_BTEDTBONPAIEMENT.OnClickReduire(Sender : Tobject);
begin
  OptionEdition.DeuxPages  := ChkReduire.checked;
  OptionEdition1.DeuxPages := ChkReduire.checked;
  OptionEdition2.DeuxPages := ChkReduire.checked;
  OptionEdition3.DeuxPages := ChkReduire.checked;
end;

procedure TOF_BTEDTBONPAIEMENT.OnClickEditJust(Sender : Tobject);
begin

  GRB_FRAIS.visible := ChkEditjust.checked

end;


Procedure TOF_BTEDTBONPAIEMENT.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
  inherited;

  Aff :=THEdit(GetControl('GP_AFFAIRE'));
  Aff0:=THEdit(GetControl('AFFAIRE0')) ;
  Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
  Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
  Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
  Aff4:=THEdit(GetControl('GP_AVENANT'))  ;

END ;

procedure TOF_BTEDTBONPAIEMENT.OnClose;
begin
  //
  FreeAndNil(OptionEdition);
  FreeAndNil(OptionEdition1);
  inherited;
end;

procedure TOF_BTEDTBONPAIEMENT.AjouteLesParamsEdt(TobEDT: TOB;OptEdt: TOptionEdition);
begin
  TobEDT.AddChampSupValeur('TIP',    OptEdt.Tip);
  TobEDT.AddChampSupValeur('NATURE', OptEdt.Nature);
  TobEDT.AddChampSupValeur('MODELE', OptEdt.Modele);
  TobEDT.AddChampSupValeur('TITRE',  OptEdt.Titre);
end;

procedure TOF_BTEDTBONPAIEMENT.AddlesChamps(TOBEdt, TOBSrc: TOB);
var Indice : integer;
begin
  Indice := 1000;
  while TOBSrc.GetNomChamp(Indice) <> '' do
  begin
    TOBEdt.AddChampSupValeur(TOBSrc.GetNomChamp(Indice),TOBSrc.GetValeur(Indice) );
    inc(indice);
  end;
end;

procedure TOF_BTEDTBONPAIEMENT.PrepareTOBEDit(TOBEc, TobedtC, TOBEdtS: TOB);
var iind,Jind : integer;
    TOBEdt1 : TOB;
begin

  if TOBEC.Detail[0].detail.count = 0 then exit;

  For iInd := 0 to TOBEC.Detail.count - 1 do
  begin
    IF (TOBEC.detail[IInd].GetString('TYPEPAIE')='001') or (TOBEC.detail[IInd].GetString('TYPEPAIE')='') then
    begin
      for jInd := 0 to TOBEC.detail[IInd].detail.count -1 do
      begin
        TobEdt1 := TOB.Create ('UNE LIGNE',TOBEdtS,-1);
        AddlesChamps (TOBEdt1,TOBEC.detail[IInd]);
        AddlesChamps (TOBEdt1,TOBEC.detail[IInd].detail[jind]);
      end;
    end else
    begin
      for jInd := 0 to TOBEC.detail[IInd].detail.count -1 do
      begin
        TobEdt1 := TOB.Create ('UNE LIGNE',TOBEdtC,-1);
        AddlesChamps (TOBEdt1,TOBEC.detail[IInd]);
        AddlesChamps (TOBEdt1,TOBEC.detail[IInd].detail[jind]);
      end;
    end;
  end;

end;

procedure TOF_BTEDTBONPAIEMENT.PrepareTOBRecap(TOBEc, TOBedtR: TOB);
var IInd,Jind : integer;
    TOBEdt1 : TOB;
begin
  if TOBEc.detail.count = 0 then exit;

  For iInd := 0 to TOBEC.Detail.count - 1 do
  begin
    for jInd := 0 to TOBEC.detail[IInd].detail.count -1 do
    begin
      TobEdt1 := TOB.Create ('UNE LIGNE',TOBEdtR,-1);
      AddlesChamps (TOBEdt1,TOBEC.detail[IInd]);
      AddlesChamps (TOBEdt1,TOBEC.detail[IInd].detail[jind]);
    end;
  end;

end;

procedure TOF_BTEDTBONPAIEMENT.EditeTout;
var TOBedit : TOB;
    iInd : integer;
    OptEdt : TOptionEdition;
begin

  OptEdt := TOptionEdition.create;
  if TOBEditions.detail.count = 0 then exit;

  StartPdfBatch;

  For iInd := 0 to TobEditions.Detail.count - 1 do
  begin
    TobEdit  := TobEditions.detail[iInd];
    if TOBEDit.detail.count = 0 then continue;
    OptEdt.Tip     := TobEdit.GetString('TIP');
    OptEdt.Nature  := TobEdit.GetString('NATURE');
    OptEdt.Modele  := TobEdit.GetString('MODELE');
    OptEdt.Titre   := TobEdit.GetString('TITRE');
    OptEdt.Apercu  := ChkApercu.Checked;
    OptEdt.DeuxPages := ChkReduire.Checked;
    OptEdt.Spages :=  TPageControl(Ecran.FindComponent('Pages'));

    if OptEdt.LanceImpression('', TobEdit) < 0 then V_PGI.IoError:=oeUnknown;
  end;

  CancelPDFBatch ;

  PreviewPDFFile('',GetMultiPDFPath,True);
  OptEdt.free;
end;



Initialization
  registerclasses ( [ TOF_BTEDTBONPAIEMENT ] ) ;
end.
