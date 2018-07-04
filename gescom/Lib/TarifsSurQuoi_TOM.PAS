{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/06/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TarifsSurQuoi (TarifsSurQuoi)
Mots clefs ... : TOM;TarifsSurQuoi
*****************************************************************}
Unit TarifsSurQuoi_TOM ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  HTB97,
  {$IFNDEF EAGLCLIENT}
    db,
    dbTables,
    Fe_Main,
    Fiche,
    FichList,
  {$ELSE}
    MainEagl,
    eFiche,
    eFichList,
  {$ENDIF}

  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  wTom,
  SaisieList,
  uTableFiltre,
  Menus,
  hPanel;

Type
  TOM_yTarifsSurQuoi = Class (twTOM)
  procedure OnNewRecord                ; override ;
  procedure OnDeleteRecord             ; override ;
  procedure OnUpdateRecord             ; override ;
  procedure OnAfterUpdateRecord        ; override ;
  procedure OnLoadRecord               ; override ;
  procedure OnChangeField ( F: TField) ; override ;
  procedure OnArgument ( S: String )   ; override ;
  procedure OnClose                    ; override ;
  procedure OnCancelRecord             ; override ;
  private
    iIdentifiantYTA  : integer;
    iDeIdentifiantYTQ: integer;

    sMethodeDeSaisie : string;
    sAction, sDroit  : string;

    sDevise          : string;
    sFonctionnalite  : string;
    sNatureAuxi      : string;
    sTiers           : string;
    sFamilleTiers    : string;
    sAffaire         : string;
    sFamilleAffaire  : string;
    sItineraire      : string;

    TF        : TTAbleFiltre;
    iPartiel  : integer;
    iSensTaux : integer;

    TobYFO : tob;

    LBLIBELLERESSOURCE  : tHLabel;
    LBLIBELLEARTICLE    : tHLabel;
    LBPRIXPOURQTE       : tHLabel;
    LBARRONDIPRIXHT     : tHLabel;
    LBARRONDIPRIXTTC    : tHLabel;
    LBQUALIFUNITEVTE    : tHLabel;

    GBTarifs            : tGroupBox;
    GBContexte          : tGroupBox;
    GBParametres        : tGroupBox;

    BAccesFourchette    : tToolbarButton97;
    BAccesConditions    : tToolbarButton97;

    YTQ_CIRCUIT : tHEdit;
    YTQ_PHASE   : tHEdit;

    { Click sur le menu du bouton bAccesFourchette }
    procedure MnSaisieGrille_OnClick(Sender: TObject);
    procedure MnSaisieOnglet_OnClick(Sender: TObject);
    procedure MnSaisieMixte_OnClick(Sender: TObject);

    { Affichage }
    procedure SetSayOuGet(sControl: string; lConditionSay, lClef: boolean );
    function  ControleAffichagePrixBrutPrixNet : boolean;
    procedure ControleAffichageRessourceTarifRessource;
    procedure ControleAffichageArticleTarifArticle;
    procedure ControleAffichageItineraireCircuitPhase;
    procedure ControleAffichageTranche;
    { Bouton }
    procedure BACCESFOURCHETTE_OnClick(Sender: TOBject);
    procedure BACCESCONDITIONS_OnClick(Sender: TOBject);

    procedure DoAfterLoadGrid( Sender: TObject );

  protected
    { Control Field }
    procedure ControlField(FieldName: string);
    procedure CalculField(FieldName: string);
    function  RecordIsValid: boolean; override;
  end ;

Implementation

uses
    wCommuns
   ,UtilArticle
   ,Tarifs
   ,TarifCond
   ,ParamSoc
   ,wJetons
   ,UtilRessource
   {$IFDEF GPAO}
   ,wItineraire
   {$ENDIF GPAO}
   ;

Const
  sFournisseur  = 'FOU';
  sOu_L = 'L';
  sOu_P = 'P';
  sQuoi_Q   = 'Q';

  sTarifFournisseur       = '101';   //Fonctionnalité : tarif fournisseur
  sCommissionFournisseur  = '102';   //Fonctionnalité : commissionnement fournisseur
  sIndirectFournisseur    = '103';   //Fonctionnalité : indirect fournisseur

  sTarifClient            = '201';   //Fonctionnalité : tarif client
  sCommissionClient       = '202';   //Fonctionnalité : commissionnement client
  sIndirectClient         = '203';   //Fonctionnalité : indirect client

  sTarifSousTraitantAchat = '301';   //Fonctionnalité : tarif sous-traitant d'achat

  sTarifSousTraitantPhase = '401';   //Fonctionnalité : tarif sous-traitant de phase

	{ Tableau des libellés de messages pour les ressources }
	TexteMessage: array[1..12] of string =
      ( {1}  'Le code article est inconnu',
			  {2}  'Le code article et le code famille tarifs d''article ne doivent pas être renseignés simultanément',
	      {3}  'Cette clef naturelle existe déjà dans cette table',
        {4}  'Le code ressource est inconnu',
			  {5}  'Le code ressource et le code famille tarifs de ressource ne doivent pas être renseignés simultanément',
        {6}  'Le code circuit est inconnu',
        {7}  'Le code circuit est inconnu pour l''itinéraire de l''article',
        {8}  'Le code phase est inconnu pour le circuit',
        { 9}  'Le code traitement est inconnu',
        {10}  'Il est incohérent de vouloir rechercher à la pièce tout en sommant à la ligne',
        {11}  'Le taux de remise ne doit pas dépasser 100%',
        {12}  'Incohérence entre le paramètre "Par fourchette" et l''écran où les données ont été saisies'
      );
	ErrArticle                   =  1;
	ErrArticleEtTarifArticle     =  2;
  ErrClefNaturelle             =  3;
	ErrRessource                 =  4;
	ErrRessourceEtTarifRessource =  5;
	ErrCircuit                   =  6;
	ErrCircuitDeItineraire       =  7;
	ErrPhaseduCircuit            =  8;
	ErrTraitement                =  9;
  ErrSommerRechercher          = 10;
  ErrRemise100                 = 11;
  ErrTrancheOK                 = 12;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnNewRecord ;
begin
  { Pour la duplication }
  if ecran <> nil then
  begin
    iDeIdentifiantYTQ := GetInteger('YTQ_IDENTIFIANT');
  end;

  Inherited ;
  //Appel depuis la fiche
  if (ecran <> nil) then
  begin
    SetInteger('YTQ_IDENTIFIANTYTA',iIdentifiantYTA);
    //En création initilisation des données principales qui peuvent l'être
    if IKC='C' then
    begin
      SetString('YTQ_RECHERCHEROU',sOu_L);
      SetString('YTQ_SOMMEROU'    ,sOu_L);
      SetString('YTQ_FOURCHETTEDE',sQuoi_Q);
    end;
    //Le controle de la clef se fait avant l'Update Record
    BAccesFourchette.Enabled := False;
  end;
//SetString('YTQ_LIBELLETARIF', SetTarifLibelle(sFonctionnalite, sTiers, sFamilleTiers, sAffaire, sFamilleAffaire, GetString('YTQ_CODEARTICLE'), GetString('YTQ_TARIFARTICLE'), GetString('YTQ_RESSOURCE'), GetString('YTQ_TARIFRESSOURCE')));
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnDeleteRecord ;
begin
  Inherited ;
  wDeleteTable('yTarifsFourchette', 'YTF_IDENTIFIANTYTA = "'+IntToStr(GetInteger('YTQ_IDENTIFIANTYTA'))+'"'+' AND '+'YTF_IDENTIFIANTYTQ = "'+IntToStr(GetInteger('YTQ_IDENTIFIANT'))+'"');
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnUpdateRecord ;
var
  TobYTQ: Tob;
begin
  DisableControl;
  try
    if Ecran = nil then
      TobYTQ := fTob
    else
    begin
    {$IFDEF EAGLCLIENT}
      TobYTQ := TFFiche(Ecran).QFiche;
    {$ELSE}
      TobYTQ := Tob.Create(FTableName, nil, -1);
      wMakeTobFromDS(DS, TobYTQ);
    {$ENDIF}
    end;

    //Calcul du Poids du contexte
    SetInteger('YTQ_POIDSRECHERCHE',CalculPoidsSurQuoi(TobYTQ,iPartiel));

    {$IFNDEF EAGLCIENT}
      if Assigned(Ecran) and Assigned(TobYTQ) then
        TobYTQ.Free;
    {$ENDIF}

    SetString('YTQ_ARTICLE',wGetArticleFromCodeArticle(GetString('YTQ_CODEARTICLE')));
    if (iPartiel<>0) and (not wExistGA(GetString('YTQ_ARTICLE'))) then
      SetString('YTQ_ARTICLE',GetString('YTQ_CODEARTICLE'));

    if RecordIsValid then
    begin
    end;

  finally
    EnableControl;
  end;

  Inherited ;
  if (Ecran<>nil) then BAccesFourchette.Enabled := GetBoolean('YTQ_PARTRANCHE');
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnAfterUpdateRecord ;
var
  iDeIdentifiantYTA, iAIdentifiantYTA, iAIdentifiantYTQ : integer;
begin
  Inherited ;
  if (iDeIdentifiantYTQ > 0) and (Getboolean('YTQ_TRANCHE')) and (PgiAsk('Confirmez-vous la duplication complète du système tarifaire ?','') = mrYes) then
  begin
    { Duplication de YTF }
    iDeIdentifiantYTA := GetInteger('YTQ_IDENTIFIANTYTA');   //Identifiant du A Qui d'origine équivalent à l'identifiant du A Qui Destinataire qui sont tous deux identiques
    iAIdentifiantYTA  := GetInteger('YTQ_IDENTIFIANTYTA');   //Identifiant du A Qui de destination
    iAIdentifiantYTQ := GetInteger('YTQ_IDENTIFIANT');

    DuplicationYTF(sFonctionnalite, iDeIdentifiantYTA , iDeIdentifiantYTQ, iAIdentifiantYTA, iAIdentifiantYTQ, 1, sDevise);
  end;
  iDeIdentifiantYTQ:=0;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnLoadRecord ;
begin
  Inherited ;
  tTabSheet(GetControl('TSJAMAISAFFICHE1')).TabVisible := False;
  tTabSheet(GetControl('TSJAMAISAFFICHE2')).TabVisible := False;
  tTabSheet(GetControl('TSJAMAISAFFICHE3')).TabVisible := False;
  ControleAffichagePrixBrutPrixNet;
  ControleAffichageRessourceTarifRessource;
  ControleAffichageArticleTarifArticle;
  ControleAffichageItineraireCircuitPhase;
  ControleAffichageTranche;
  BAccesFourchette.Enabled := GetBoolean('YTQ_PARTRANCHE');
end;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnArgument ( S: String ) ;
var
  lSaisieEnGrille   : boolean;
  aAqui             : MyArrayValue;
  iRangPanel, iHauteurDisponible : integer;

  procedure ActivePanel(sNomPanel: string; var iRangPanel: integer; var iHauteurDisponible : integer; iHauteurPanel : integer);
  begin
    //Tabsheet TSCaractérisitiques complet on active au Tabsheet Page2
    if (iHauteurDisponible < iHauteurPanel) then
    begin
      if (not tTabSheet(GetControl('PAGE2')).TabVisible) then
      begin
        iHauteurDisponible := 156; //Hauteur du TabSheet
        iRangPanel         :=   3; //Premier Panel du Tabsheet complément
        tTabSheet(GetControl('PAGE2')).TabVisible := True;
      end
      else
      begin
        iHauteurDisponible := 156; //Hauteur du TabSheet
        iRangPanel         :=   5; //Premier Panel du Tabsheet complément
        tTabSheet(GetControl('TSLIGNE')).TabVisible := True;
      end;
    end;

    thPanel(GetControl(sNomPanel)).Parent                 := tHPanel(GetControl('PN'+IntToStr(iRangPanel)));
    thPanel(GetControl('PN'+IntToStr(iRangPanel))).Height := iHauteurPanel;

    iRangPanel := iRangPanel+1;
    iHauteurDisponible := iHauteurDisponible - iHauteurPanel;
  end;

begin
  Inherited ;
  iIdentifiantYTA   := ValeurI(GetArgumentValue(S,'YTA_IDENTIFIANT'));
  sMethodeDeSaisie  := GetArgumentValue(S,'METHODEDESAISIE');
  sAction           := GetArgumentValue(S,'ACTION');
  sDroit            := GetArgumentValue(S,'DROIT');

  aAqui := WGetSqlFieldsValues(['YTA_FONCTIONNALITE', 'YTA_DEVISE', 'YTA_TIERS', 'YTA_TARIFTIERS', 'YTA_AFFAIRE', 'YTA_TARIFAFFAIRE'], 'YTARIFSAQUI', 'YTA_IDENTIFIANT="'+IntToStr(iIdentifiantYTA)+'"');
  if (aAQui<>nil) then
  begin
    sFonctionnalite   := aAQui[0];
    sDevise           := aAQui[1];
    sTiers            := aAQui[2];
    sFamilleTiers     := aAqui[3];
    sAffaire          := aAQui[4];
    sFamilleAffaire   := aAqui[5];
    sNatureAuxi       := NatureAuxiliaire(sFonctionnalite);
  end;

  Ecran.Caption := RechDom('YFONCTIONNALITES',sFonctionnalite,False)+TraduireMemoire(' : Sur Quoi');

  if (Ecran<>nil) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    lSaisieEnGrille := (sMethodeDeSaisie='GRILLE');
    TF.SaisieMixte  := (sMethodeDeSaisie='MIXTE');
    TF.SaisieEnGrid := ((sMethodeDeSaisie='MIXTE') or (sMethodeDeSaisie='GRILLE'));
    TFSaisieList(Ecran).PCPied.Visible        := not lSaisieEnGrille;
    TFSaisieList(Ecran).PanPied.visible       := not lSaisieEnGrille;
    TFSaisieList(Ecran).SplitMessages.Visible := not lSaisieEnGrille;
    TF.OnAfterLoadGrid := DoAfterLoadGrid;
  end;

  LBLIBELLERESSOURCE  := thLabel(GetControl('LBLIBELLERESSOURCE'));
  LBLIBELLEARTICLE    := thLabel(GetControl('LBLIBELLEARTICLE'));
  LBPRIXPOURQTE       := thLabel(GetControl('LBPRIXPOURQTE')   );
  LBARRONDIPRIXHT     := thLabel(GetControl('LBARRONDIPRIXHT' ));
  LBARRONDIPRIXTTC    := thLabel(GetControl('LBARRONDIPRIXTTC'));
  LBQUALIFUNITEVTE    := thLabel(GetControl('LBQUALIFUNITEVTE'));

  GBTarifs            := tGroupBox(GetControl('GBTARIFS'));
  GBContexte          := tGroupBox(GetControl('GBCONTEXTE'));
  GBParametres        := tGroupBox(GetControl('GBPARAMETRES'));

  YTQ_CIRCUIT         := tHEdit(GetControl('YTQ_CIRCUIT'));
  YTQ_PHASE           := tHEdit(GetControl('YTQ_PHASE'));

  BAccesFourchette    := tToolbarButton97(GetControl('BACCESFOURCHETTE'));
  BAccesConditions    := tToolbarButton97(GetControl('BACCESCONDITIONS'));

  { Menu associé au bouton BAccesFourchette }
  if (GetControl('MNSAISIEGRILLE')  <>nil) then TMenuItem(GetControl('MNSAISIEGRILLE')).OnClick := MnSaisieGrille_OnClick;
  if (GetControl('MNSAISIEONGLET')  <>nil) then TMenuItem(GetControl('MNSAISIEONGLET')).OnClick := MnSaisieOnglet_OnClick;
  if (GetControl('MNSAISIEMIXTE')   <>nil) then TMenuItem(GetControl('MNSAISIEMIXTE') ).OnClick := MnSaisieMixte_OnClick;
  if (GetControl('BACCESCONDITIONS')<>nil) then TToolBarButton97(GetControl('BACCESCONDITIONS')).OnClick := BACCESCONDITIONS_OnClick ;
  if (GetControl('BACCESFOURCHETTE')<>nil) then TToolBarButton97(GetControl('BACCESFOURCHETTE')).OnClick := BACCESFOURCHETTE_OnClick ;

  //Recherche d'un tarif sur code article partiel
  iPartiel  := GetParamSoc('SO_TARIFPARTIEL');

  //Les taux sont-ils à retrancher ( remise ) ou à augmenter ( taux commissionnement par exemple )
  if (sFonctionnalite=sCommissionClient) or (sFonctionnalite=sCommissionFournisseur) then
    iSensTaux := +1
  else
    iSensTaux := -1;

  SetSayOuGet('YTQ_BLOCNOTE',True,False);

  //Fixer les libellés par rapport à la fonctionnalité
  if (GetControl('TYTQ_FORFAIT')       <>nil) then tHLabel(GetControl('TYTQ_FORFAIT')      ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'FORFAIT'      );
  if (GetControl('TYTQ_COUTLANCEMENT') <>nil) then tHLabel(GetControl('TYTQ_COUTLANCEMENT')).Caption := LibelleParFonctionnalite(sFonctionnalite, 'COUTLANCEMENT');
  if (GetControl('TYTQ_REMISEMONTANT') <>nil) then tHLabel(GetControl('TYTQ_REMISEMONTANT')).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISEMONTANT');
  if (GetControl('TYTQ_PRIXBRUT')      <>nil) then tHLabel(GetControl('TYTQ_PRIXBRUT')     ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXBRUT'     );
  if (GetControl('TYTQ_REMISE1')       <>nil) then tHLabel(GetControl('TYTQ_REMISE1')      ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISE123'    );
  if (GetControl('TYTQ_PRIXNET')       <>nil) then tHLabel(GetControl('TYTQ_PRIXNET')      ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXNET'      );
  if (GetControl('TYTQ_CONDFORFAIT')   <>nil) then tHLabel(GetControl('TYTQ_CONDFORFAIT')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'FORFAIT'      );
  if (GetControl('TYTQ_CONDCOUTLANC')  <>nil) then tHLabel(GetControl('TYTQ_CONDCOUTLANC') ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'COUTLANCEMENT');
  if (GetControl('TYTQ_CONDREMISEM')   <>nil) then tHLabel(GetControl('TYTQ_CONDREMISEM')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISEMONTANT');
  if (GetControl('TYTQ_CONDPRIXBRUT')  <>nil) then tHLabel(GetControl('TYTQ_CONDPRIXBRUT') ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXBRUT'     );
  if (GetControl('TYTQ_CONDREMISE1')   <>nil) then tHLabel(GetControl('TYTQ_CONDREMISE1')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'REMISE123'    );
  if (GetControl('TYTQ_CONDPRIXNET')   <>nil) then tHLabel(GetControl('TYTQ_CONDPRIXNET')  ).Caption := LibelleParFonctionnalite(sFonctionnalite, 'PRIXNET'      );

  //Donner la possibiliter de saisir des conditions d'application
  BAccesConditions.Visible := GetParamSoc('SO_TARIFCONDAPPL');

  BAccesFourchette.Enabled := (TF.RecCount<>0);

  //Fixe l'écran de saisi en fonction des paramètres de gestion des informations de la fonctionnalite
  TobYFO := Tob.Create('_YFO_',nil,-1);
  GetTobYFO(sFonctionnalite, TobYFO);
  //Ces deux TabSheets regroupent toutes les données disponibles mais sont invisibles
  tTabSheet(GetControl('TSJAMAISAFFICHE1')).TabVisible := False;
  tTabSheet(GetControl('TSJAMAISAFFICHE2')).TabVisible := False;
  tTabSheet(GetControl('TSJAMAISAFFICHE3')).TabVisible := False;
  tTabSheet(GetControl('PAGE2')).TabVisible            := False;
  //Sur ces TabSheets de saisis transfert des données saisissables : paramétrées comme telles dans la fonctionnalité demandée
  iRangPanel:=1; iHauteurDisponible := 156;
  if (TobYFO.Detail[0].GetValue('YFO_OKARTICLE')   =wTrue) then ActivePanel('PNARTICLE'       , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKRESSOURCE') =wTrue) then ActivePanel('PNRESSOURCE'     , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKCPT')       =wTrue) then ActivePanel('PNSOUSTRAITANCE' , iRangPanel, iHauteurDisponible, 104);
  if (TobYFO.Detail[0].GetValue('YFO_OKMARQUE')    =wTrue) then ActivePanel('PNMARQUE'        , iRangPanel, iHauteurDisponible,  52);
  if (TobYFO.Detail[0].GetValue('YFO_OKYTQLIGNE')  =wTrue) then ActivePanel('PNCHR'           , iRangPanel, iHauteurDisponible, 104);
  if (TobYFO.Detail[0].GetValue('YFO_OKYTQLIGNE')  =wTrue) then ActivePanel('PNAFFAIRE'       , iRangPanel, iHauteurDisponible,  52);
  //Activation du premier TabSheet de saisie
  tPageControl(GetControl('PCPied')).ActivePage := tTabSheet(GetControl('TSCONDITIONS'));
  tPageControl(GetControl('PCPied')).TabWidth := 100;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnClose ;
begin
  Inherited ;
  if assigned(TobYFO) then TobYFO.Free;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.OnCancelRecord ;
begin
  Inherited ;
  BAccesFourchette.Enabled := GetBoolean('YTQ_PARTRANCHE');
  ControleAffichagePrixBrutPrixNet;
  ControleAffichageRessourceTarifRessource;
  ControleAffichageArticleTarifArticle;
  ControleAffichageItineraireCircuitPhase;
  ControleAffichageTranche;
end ;

{--------------------------------------------------------------------------------
   Data Change Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.CalculField(FieldName: string);
var
  aArticle, aRessource : MyArrayValue;
begin
  aArticle   := nil;
  aRessource := nil;
//  inherited;
  if      (FieldName='YTQ_CODEARTICLE') then
  begin
    aArticle := wGetSqlFieldsValues(['GA_LIBELLE', 'GA_PRIXPOURQTE', 'GA_ARRONDIPRIX', 'GA_ARRONDIPRIXTTC', 'GA_QUALIFUNITEVTE'], 'ARTICLE', 'GA_CODEARTICLE="'+GetString(FieldName)+'"');
    if (aArticle<>nil) then
    begin
      LBLIBELLEARTICLE.Caption   := copy(aArticle[0],1,60);
      LBPRIXPOURQTE.Caption      := TraduireMemoire('Prix indiqués pour une quantité de ')      +IntToStr(aArticle[1]);
      LBARRONDIPRIXHT.Caption    := TraduireMemoire('Méthode d''arrondi pour les prix H.T. ')   +aArticle[2];
      LBARRONDIPRIXTTC.Caption   := TraduireMemoire('Méthode d''arrondi pour les prix T.T.C. ') +aArticle[3];
      LBQUALIFUNITEVTE.Caption   := TraduireMemoire('Unité de Facturation ')                    +aArticle[4];
    end
    else
    begin
      LBLIBELLEARTICLE.Caption   := '';
      LBPRIXPOURQTE.Caption      := '';
      LBARRONDIPRIXHT.Caption    := '';
      LBARRONDIPRIXTTC.Caption   := '';
      LBQUALIFUNITEVTE.Caption   := '';
    end;
  end
  else if (FieldName='YTQ_RESSOURCE') then
  begin
    aRessource := wGetSqlFieldsValues(['ARS_LIBELLE'], 'RESSOURCE', 'ARS_RESSOURCE="'+GetString(FieldName)+'"');
    if (aRessource<>nil) then
      LBLIBELLERESSOURCE.Caption   := aRessource[0]
    else
      LBLIBELLERESSOURCE.Caption   := '';
  end
  else if (FieldName='YTQ_PARTRANCHE') then
  begin
    (* // à remettre dès le nouvel AGL 555
    if (TF<>nil) and GBTarifs.Visible then
      TF.ChangeVisible( GBTarifs );
    *)
    GBTarifs.Visible         := ( not GetBoolean('YTQ_PARTRANCHE'));
    GBTarifs.Enabled         := ( not GetBoolean('YTQ_PARTRANCHE')) and  (not (sAction='CONSULTATION'));
    SetSayOuGet('YTQ_FORFAIT'      , GetBoolean('YTQ_PARTRANCHE'), False);
    SetSayOuGet('YTQ_COUTLANCEMENT', GetBoolean('YTQ_PARTRANCHE'), False);
    SetSayOuGet('YTQ_REMISEMONTANT', GetBoolean('YTQ_PARTRANCHE'), False);
    SetSayOuGet('YTQ_PRIXBRUT'     , GetBoolean('YTQ_PARTRANCHE'), False);
    SetSayOuGet('YTQ_REMISE1'      , GetBoolean('YTQ_PARTRANCHE'), False);
    SetSayOuGet('YTQ_REMISE2'      , GetBoolean('YTQ_PARTRANCHE'), False);
    SetSayOuGet('YTQ_REMISE3'      , GetBoolean('YTQ_PARTRANCHE'), False);
    SetSayOuGet('YTQ_PRIXNET'      , GetBoolean('YTQ_PARTRANCHE'), False);

    GBContexte.Visible       := ( not GetBoolean('YTQ_PARTRANCHE'));
    GBContexte.Enabled       := ( not GetBoolean('YTQ_PARTRANCHE')) and  (not (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDFORFAIT' , 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDCOUTLANC', 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDREMISEM' , 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDPRIXBRUT', 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDREMISE1' , 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDREMISE2' , 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDREMISE3' , 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));
    SetControlProperty('YTQ_CONDPRIXNET' , 'Enabled', not GetBoolean('YTQ_PARTRANCHE') or (sAction='CONSULTATION'));

    GBParametres.Visible       := ( GetBoolean('YTQ_PARTRANCHE'));
    GBParametres.Enabled       := ( GetBoolean('YTQ_PARTRANCHE')) and  (not (sAction='CONSULTATION'));

    BAccesFourchette.Enabled := GetBoolean('YTQ_PARTRANCHE') and (TF.State in [dsBrowse]);
  end
  ;
end;

{--------------------------------------------------------------------------------
   Validation Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.ControlField(FieldName: string);
begin
 	DisableControl;
  try
    if      (FieldName = 'YTQ_FORFAIT')         then SetBoolean('YTQ_CONDFORFAIT' ,GetDouble('YTQ_FORFAIT')       <>0.0)
    else if (FieldName = 'YTQ_COUTLANCEMENT')   then SetBoolean('YTQ_CONDCOUTLANC',GetDouble('YTQ_COUTLANCEMENT') <>0.0)
    else if (FieldName = 'YTQ_REMISEMONTANT')   then SetBoolean('YTQ_CONDREMISEM' ,GetDouble('YTQ_REMISEMONTANT') <>0.0)
    else if (FieldName = 'YTQ_PRIXBRUT')        then SetBoolean('YTQ_CONDPRIXBRUT',GetDouble('YTQ_PRIXBRUT')      <>0.0)
    else if (FieldName = 'YTQ_REMISE1')         then SetBoolean('YTQ_CONDREMISE1' ,GetDouble('YTQ_REMISE1')       <>0.0)
    else if (FieldName = 'YTQ_REMISE2')         then SetBoolean('YTQ_CONDREMISE2' ,GetDouble('YTQ_REMISE2')       <>0.0)
    else if (FieldName = 'YTQ_REMISE3')         then SetBoolean('YTQ_CONDREMISE3' ,GetDouble('YTQ_REMISE3')       <>0.0)
    else if (FieldName = 'YTQ_PRIXNET')         then SetBoolean('YTQ_CONDPRIXNET' ,GetDouble('YTQ_PRIXNET')       <>0.0)
    ;

    if (GetParamSoc('SO_TARIFLIBAUTO')) and (Pos(FieldName,'YTQ_CODEARTICLE/YTQ_TARIFARTICLE/YTQ_RESSOURCE/YTQ_TARIFRESSOURCE')>0) then
    begin
      SetString('YTQ_LIBELLETARIF', SetTarifLibelle(sFonctionnalite, sTiers, sFamilleTiers, sAffaire, sFamilleAffaire, GetString('YTQ_CODEARTICLE'), GetString('YTQ_TARIFARTICLE'), GetString('YTQ_RESSOURCE'), GetString('YTQ_TARIFRESSOURCE')));
    end;

    if (Pos(FieldName,'YTQ_PRIXBRUT/YTQ_REMISE1/YTQ_REMISE2/YTQ_REMISE3/YTQ_PRIXNET')>0) then
    begin
      if (Pos(FieldName,'YTQ_REMISE1')>0) and (GetDouble('YTQ_REMISE1')>100) then LastError := ErrRemise100;
      if (Pos(FieldName,'YTQ_REMISE2')>0) and (GetDouble('YTQ_REMISE2')>100) then LastError := ErrRemise100;
      if (Pos(FieldName,'YTQ_REMISE3')>0) and (GetDouble('YTQ_REMISE3')>100) then LastError := ErrRemise100;
      if ControleAffichagePrixBrutPrixNet then
        SetDouble('YTQ_PRIXNET',CalculPrixNet( iif(GetBoolean('YTQ_CONDPRIXBRUT'),GetDouble('YTQ_PRIXBRUT'),0)
                                              ,iif(GetBoolean('YTQ_CONDREMISE1') ,GetDouble('YTQ_REMISE1') ,0)
                                              ,iif(GetBoolean('YTQ_CONDREMISE2') ,GetDouble('YTQ_REMISE2') ,0)
                                              ,iif(GetBoolean('YTQ_CONDREMISE3') ,GetDouble('YTQ_REMISE3') ,0)
                                              ,iSensTaux,sDevise,sNatureAuxi,'P','P'));
    end
    else if (Pos(FieldName,'YTQ_CODEARTICLE/YTQ_TARIFARTICLE')>0) then
    begin
      ControleAffichageArticleTarifArticle;
      SetString('YTQ_ARTICLE',wGetArticleFromCodeArticle(GetString('YTQ_CODEARTICLE')));
      if (iPartiel<>0) and (not wExistGA(GetString('YTQ_ARTICLE'))) then
         SetString('YTQ_ARTICLE',GetString('YTQ_CODEARTICLE'));
      ControleAffichageItineraireCircuitPhase;
    end
    else if (Pos(FieldName,'YTQ_CIRCUIT')>0) then
    begin
      ControleAffichageItineraireCircuitPhase;
    end
    else if (Pos(FieldName,'YTQ_PARTRANCHE')>0) then
    begin
      ControleAffichageTranche;
    end
    else if (Pos(FieldName,'YTQ_SOMMEROU/YTQ_RECHERCHEROU')>0)  then
    begin
      if (GetField('YTQ_SOMMEROU')=sOu_L) and (GetField('YTQ_RECHERCHEROU')=sOu_P) then
        LastError := ErrSommerRechercher;
    end
    ;
  finally
    EnableControl;
  end;

//  inherited;
  BAccesFourchette.Enabled := False;

  if LastError <> 0 then
    LastErrorMsg := TexteMessage[LastError];
end;

{--------------------------------------------------------------------------------
   Validation Global
---------------------------------------------------------------------------------}
function TOM_yTarifsSurQuoi.RecordIsValid: boolean;
var
  sClefNaturelle : string;
  lTarifCondAppl : boolean;
  lTrancheOK  : boolean;
begin
  // controle de la clef naturelle
  sClefNaturelle :=        '(YTQ_IDENTIFIANTYTA="' +IntToStr(GetInteger('YTQ_IDENTIFIANTYTA'))+'") '+
                       'and (YTQ_RESSOURCE="'      +GetString('YTQ_RESSOURCE')                +'") '+
                       'and (YTQ_TARIFRESSOURCE="' +GetString('YTQ_TARIFRESSOURCE')           +'") '+
                       'and (YTQ_ARTICLE="'        +GetString('YTQ_ARTICLE')                  +'") '+
                       'and (YTQ_TARIFARTICLE="'   +GetString('YTQ_TARIFARTICLE')             +'") '+
                       'and (YTQ_CIRCUIT="'        +GetString('YTQ_CIRCUIT')                  +'") '+
                       'and (YTQ_PHASE="'          +GetString('YTQ_PHASE')                    +'") '+
                       'and (YTQ_TRAITEMENT="'     +GetString('YTQ_TRAITEMENT')               +'") '+
                       'and (YTQ_MARQUE="'         +GetString('YTQ_MARQUE')                   +'") '+
                       '';
  if (IKC='M') then
    sClefNaturelle := sClefNaturelle + ' and (YTQ_IDENTIFIANT<>'+ IntToStr(GetInteger('YTQ_IDENTIFIANT')) +')';

  lTarifCondAppl := GetParamSoc('SO_TARIFCONDAPPL');
  if (ExistYTQ( sClefNaturelle )) and ((not lTarifCondAppl) or ( not (PgiAsk('Confirmez-vous l''enregistrement de cette fiche ?','ATTENTION cette clef naturelle existe déjà  !') = mrYes))) then
    LastError := ErrClefNaturelle;

  lTrancheOK := True;
  if (GetBoolean('YTQ_PARTRANCHE')) and ((GetDouble('YTQ_FORFAIT')<>0.0) or (GetDouble('YTQ_COUTLANCEMENT')<>0.0) or (GetDouble('YTQ_REMISEMONTANT')<>0.0) or (GetDouble('YTQ_PRIXBRUT')<>0.0) or (GetDouble('YTQ_REMISE1')<>0.0) or (GetDouble('YTQ_REMISE2')<>0.0) or (GetDouble('YTQ_REMISE3')<>0.0) or (GetDouble('YTQ_PRIXNET')<>0.0)) then
  begin
    lTrancheOK := False;
    if (PgiAsk('Des conditions tarifaires existent : Confirmez-vous la remise à zéro de ces conditions ?','ATTENTION') = mrYes) then
    begin
      SetDouble('YTQ_FORFAIT',0.0)        ; ControlField('YTQ_FORFAIT');
      SetDouble('YTQ_COUTLANCEMENT',0.0)  ; ControlField('YTQ_COUTLANCEMENT');
      SetDouble('YTQ_REMISEMONTANT',0.0)  ; ControlField('YTQ_REMISEMONTANT');
      SetDouble('YTQ_PRIXBRUT',0.0)       ; ControlField('YTQ_PRIXBRUT');
      SetDouble('YTQ_REMISE1',0.0)        ; ControlField('YTQ_REMISE1');
      SetDouble('YTQ_REMISE2',0.0)        ; ControlField('YTQ_REMISE2');
      SetDouble('YTQ_REMISE3',0.0)        ; ControlField('YTQ_REMISE3');
      SetDouble('YTQ_PRIXNET',0.0)        ; ControlField('YTQ_PRIXNET');
      lTrancheOK := True;
    end
    else
    begin
      SetBoolean('YTQ_PARTRANCHE',False);
    end;
  end;
  if (not GetBoolean('YTQ_PARTRANCHE')) and ExistTarifsFourchette(GetInteger('YTQ_IDENTIFIANTYTA'), GetInteger('YTQ_IDENTIFIANT')) then
  begin
    lTrancheOK := False;
    if (PgiAsk('Des conditions tarifaires par fourchettes existent : Confirmez-vous la suppression de ces conditions ?','ATTENTION') = mrYes) then
    begin
      lTrancheOK := True;
      wDeleteTable('yTarifsFourchette', 'YTF_IDENTIFIANTYTA = "'+IntToStr(GetInteger('YTQ_IDENTIFIANTYTA'))+'"'+' AND '+'YTF_IDENTIFIANTYTQ = "'+IntToStr(GetInteger('YTQ_IDENTIFIANT'))+'"');
    end
    else
    begin
      SetBoolean('YTQ_PARTRANCHE',True);
    end;
  end;

   if      (GetString('YTQ_ARTICLE')        <>'') and (iPartiel=0) and (not wExistGA(GetString('YTQ_ARTICLE')))                         then LastError := ErrArticle
   else if (GetString('YTQ_TARIFARTICLE')   <>'') and (GetString('YTQ_ARTICLE')<>'')                                                    then LastError := ErrArticleEtTarifArticle
   else if (GetString('YTQ_RESSOURCE')      <>'') and (not ExisteRessource(GetString('YTQ_RESSOURCE')))                                 then LastError := ErrRessource
   else if (GetString('YTQ_TARIFRESSOURCE') <>'') and (GetString('YTQ_RESSOURCE')<>'')                                                  then LastError := ErrRessourceEtTarifRessource
   {$IFDEF GPAO}
     else if (GetString('YTQ_CIRCUIT')        <>'') and (                      not wExistCircuit(GetString('YTQ_CIRCUIT')))               then LastError := ErrCircuit
     else if (GetString('YTQ_CIRCUIT')        <>'') and (sItineraire<>'') and (not wExistCircuit(GetString('YTQ_CIRCUIT'), sItineraire))  then LastError := ErrCircuitDeItineraire
     else if (GetString('YTQ_PHASE')          <>'') and (not wExistOpeCirc(GetString('YTQ_PHASE'), GetString('YTQ_CIRCUIT')) )            then LastError := ErrPhaseduCircuit
   {$ENDIF GPAO}
  else if (not lTrancheOK) then LastError := ErrTrancheOK
   ;
   Result := LastError = 0;

   if LastError > 0 then
   begin
   	 LastErrorMsg := TexteMessage[LastError];
     Case LastError of
       ErrArticle                   : SetFocusControl('YTQ_CODEARTICLE');
       ErrArticleEtTarifArticle     : SetFocusControl('YTQ_TARIFARTICLE');
       ErrRessource                 : SetFocusControl('YTQ_RESSOURCE');
       ErrRessourceEtTarifRessource : SetFocusControl('YTQ_TARIFRESSOURCE');
     end;
   end;
end;

{-------------------------------------------------------------------------------
   Positionne les propriété Enable et Color d'un controle en fonction d'une condition de Say
--------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.SetSayOuGet(sControl: string; lConditionSay, lClef :boolean);
begin
  if lConditionSay or (sAction='CONSULTATION') or ( lClef and (TF.State<>dsInsert)) then
  begin
    SetControlProperty(sControl, 'Enabled', False);
    SetControlProperty(sControl, 'Color'  , ClBtnFace);
  end
  else
  begin
    SetControlProperty(sControl, 'Enabled', True);
    SetControlProperty(sControl, 'Color'  , ClWindow);
  end
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zone Prix Brute / Remise 1/2/3 et Prix Net
---------------------------------------------------------------------------------}
function TOM_yTarifsSurQuoi.ControleAffichagePrixBrutPrixNet : boolean;
var
   lPrixBrut, lNiLunNiLautre : boolean;
begin
  lPrixBrut      := (
                        ((GetDouble('YTQ_PRIXBRUT')<>0) and (GetBoolean('YTQ_CONDPRIXBRUT')))
                     or ((GetDouble('YTQ_REMISE1') <>0) and (GetBoolean('YTQ_CONDREMISE1') ))
                     or ((GetDouble('YTQ_REMISE2') <>0) and (GetBoolean('YTQ_CONDREMISE2') ))
                     or ((GetDouble('YTQ_REMISE3') <>0) and (GetBoolean('YTQ_CONDREMISE3') ))
                    );
  lNiLunNiLautre := (not lPrixBrut) and (GetDouble('YTQ_PRIXNET')=0);
  SetSayOuGet('YTQ_PRIXNET' , (    lPrixBrut) and (not lNiLunNiLautre), False);
  SetSayOuGet('YTQ_PRIXBRUT', (not lPrixBrut) and (not lNiLunNiLautre), False);
  SetSayOuGet('YTQ_REMISE1' , (not lPrixBrut) and (not lNiLunNiLautre), False);
  SetSayOuGet('YTQ_REMISE2' , (not lPrixBrut) and (not lNiLunNiLautre), False);
  SetSayOuGet('YTQ_REMISE3' , (not lPrixBrut) and (not lNiLunNiLautre), False);
  Result := lPrixBrut or lNiLunNiLautre;
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zones Ressource et Famille Tarif Ressource
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.ControleAffichageRessourceTarifRessource;
var
  lRessource, lNiLunNiLautre : boolean;
begin
  lRessource := (GetString('YTQ_RESSOURCE')<>'');
  lNiLunNiLautre := (not lRessource) and (GetString('YTQ_TARIFRESSOURCE')='');
  SetSayOuGet('YTQ_RESSOURCE'     , (not lRessource) and (not lNiLunNiLautre), False);
  SetSayOuGet('YTQ_TARIFRESSOURCE', (    lRessource) and (not lNiLunNiLautre), False);
//  Result := (not lRessource) and (not lNiLunNiLautre);
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zones Article et Famille Tarif Article
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.ControleAffichageArticleTarifArticle;
var
  lArticle, lNiLunNiLautre : boolean;
begin
  lArticle := (GetString('YTQ_CODEARTICLE')<>'');
  lNiLunNiLautre := (not lArticle) and (GetString('YTQ_TARIFARTICLE')='');
  SetSayOuGet('YTQ_CODEARTICLE' , (not lArticle) and (not lNiLunNiLautre), False);
  SetSayOuGet('YTQ_TARIFARTICLE', (    lArticle) and (not lNiLunNiLautre), False);
//  Result := (not lArticle) and (not lNiLunNiLautre);
end;

{--------------------------------------------------------------------------------
   Définition des DATATYPE et PLUS sur les Circuits et Phases en fonction de l'article
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.ControleAffichageItineraireCircuitPhase;
var
  sArticle, sNatureTravail : string;
  aArticleNature : MyArrayValue;
begin
  aArticleNature := nil;

  // Si tarif Sous traitant
  if (YTQ_CIRCUIT<>nil) and (YTQ_PHASE<>nil) then
  begin
    SetControlProperty('YTQ_CIRCUIT', 'Plus'     , ''       ); // Pas de filtre sur les circuit
    SetControlProperty('YTQ_PHASE'  , 'DataType' ,'QUTPHASE'); // Liste sur la table des phases

    sArticle        := wGetArticleFromCodeArticle(GetString('YTQ_CODEARTICLE'));
    sNatureTravail  := 'FAB';
    sItineraire     := '';

    // Si Tarif sur un article
    if (sArticle<>'') then
    begin
      aArticleNature := WGetSqlFieldsValues(['WAN_CODITI', 'WAN_CIRCUIT'], 'WARTNAT', 'WAN_ARTICLE = "'+sArticle+'" AND WAN_NATURETRAVAIL = "'+sNatureTravail+'"');
      // Article avec une nature de travail
      if (aArticleNature<>nil) and (aArticleNature[0]<>'') then
      begin
        sItineraire := aArticleNature[0];
        SetControlProperty('YTQ_CIRCUIT', 'DataType', 'QUTCIRCUIT');
        SetControlProperty('YTQ_CIRCUIT', 'Plus'    , 'QCI_CODITI="'+sItineraire+'"'); // Filtre sur les circuit de l'itinéraire
        SetControlProperty('YTQ_PHASE'  , 'DataType', 'WPHASEITINERAIRE'            ); // Liste sur les phases de l'itinéraire
        SetControlProperty('YTQ_PHASE'  , 'Plus'    , 'QP_CODITI="' +sItineraire+'"'); // filtre sur les phases de l'itinéraire
      end
      else
      begin
        SetControlProperty('YTQ_CIRCUIT', 'DataType', '');
        SetControlProperty('YTQ_CIRCUIT', 'Plus'    , '');
        SetControlProperty('YTQ_PHASE'  , 'DataType', '');
        SetControlProperty('YTQ_PHASE'  , 'Plus'    , '');
      end;
    end
    else
    begin
      SetControlProperty('YTQ_CIRCUIT', 'DataType', 'QUTCIRCUIT');
      SetControlProperty('YTQ_CIRCUIT', 'Plus'    , '');
      SetControlProperty('YTQ_PHASE'  , 'DataType', 'WPHASEDUCIRCUIT');
      SetControlProperty('YTQ_PHASE'  , 'Plus'    , 'QDE_CIRCUIT="'+GetString('YTQ_CIRCUIT')+'"');
    end;
  end;
end;

{--------------------------------------------------------------------------------
   Fonction du paramètre "Par fouchette" affichage des GroupBox correspondant
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.ControleAffichageTranche;
begin
  GBTarifs.Visible     := ( not GetBoolean('YTQ_PARTRANCHE'));
  GBTarifs.Enabled     := ( not GetBoolean('YTQ_PARTRANCHE')) and  (not (sAction='CONSULTATION'));
  GBContexte.Visible   := ( not GetBoolean('YTQ_PARTRANCHE'));
  GBContexte.Enabled   := ( not GetBoolean('YTQ_PARTRANCHE')) and  (not (sAction='CONSULTATION'));
  GBParametres.Visible := ( GetBoolean('YTQ_PARTRANCHE'));
  GBParametres.Enabled := ( GetBoolean('YTQ_PARTRANCHE')) and  (not (sAction='CONSULTATION'));
end;

{--------------------------------------------------------------------------------
   Saisie des conditions complémentaires pour l'application d'un tarif
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.BACCESCONDITIONS_OnClick(Sender: TOBject);
Var
  TheTobTarif : Tob;
  sConditionComplementaire : string;
begin
  TheTobTarif:=Tob.Create('_CONDITIONAPPLICATION_',nil,-1);
  try
    TheTobTarif.AddChampSupValeur('GF_CONDAPPLIC', GetString('YTQ_BLOCNOTE'), false);
    sConditionComplementaire := GetString('YTQ_BLOCNOTE');
    EntreeTarifCond(taModif, TheTobTarif, 'Article');
    if (sConditioncomplementaire<>TheTobTarif.GetValue('GF_CONDAPPLIC')) then
    begin
      if (not(TF.State in [dsInsert,dsEdit])) then
        TF.Edit;
      SetString('YTQ_BLOCNOTE',TheTobTarif.GetValue('GF_CONDAPPLIC'));
    end;
  finally
    TheTobTarif.Free;
  end;
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs fourchettes en saisie en grille
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.MnSaisieGrille_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSFOURCH_FSL','YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';'+'YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT')+';METHODEDESAISIE=GRILLE'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+sDevise);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs fourchettes en saisie en onglet
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.MnSaisieOnglet_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSFOURCH_FSL','YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';'+'YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT')+';METHODEDESAISIE=ONGLET'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+sDevise);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs fourchettes en saisie mixte
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.MnSaisieMixte_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSFOURCH_FSL','YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';'+'YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT')+';METHODEDESAISIE=MIXTE'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+sDevise);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Fourchettes selon le paramètre
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.BACCESFOURCHETTE_OnClick(Sender: TOBject);
var
  sAppelMethodeSaisie, sParamMethodeSaisie : string;
begin
  sParamMethodeSaisie := GetParamSoc('SO_TARIFSAISIE');
  if      (sParamMethodeSaisie = 'O') then sAppelMethodeSaisie := 'ONGLET'
  else if (sParamMethodeSaisie = 'G') then sAppelMethodeSaisie := 'GRILLE'
  else if (sParamMethodeSaisie = 'S') then sAppelMethodeSaisie := 'MIXTE'
  else                                     sAppelMethodeSaisie := 'MIXTE';
  AGLLanceFiche('Y','YTARIFSFOURCH_FSL','YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';'+'YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';YTF_IDENTIFIANTYTA='+GetString('YTQ_IDENTIFIANTYTA')+';YTF_IDENTIFIANTYTQ='+GetString('YTQ_IDENTIFIANT')+';METHODEDESAISIE='+sAppelMethodeSaisie+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+sDevise);
end;

{--------------------------------------------------------------------------------
   A chaque rafraichissement de la grille
---------------------------------------------------------------------------------}
procedure TOM_yTarifsSurQuoi.DoAfterLoadGrid(Sender: TObject);
begin
  if (GetControl('TSJAMAISAFFICHE1')<>nil) then tTabSheet(GetControl('TSJAMAISAFFICHE1')).TabVisible := False;
  if (GetControl('TSJAMAISAFFICHE2')<>nil) then tTabSheet(GetControl('TSJAMAISAFFICHE2')).TabVisible := False;
  if (GetControl('TSJAMAISAFFICHE3')<>nil) then tTabSheet(GetControl('TSJAMAISAFFICHE3')).TabVisible := False;
  SetControlProperty('TSSAV', 'TABVISIBLE', V_PGI.sav);
end;

Initialization
  registerclasses ( [ TOM_yTarifsSurQuoi ] ) ;
end.
