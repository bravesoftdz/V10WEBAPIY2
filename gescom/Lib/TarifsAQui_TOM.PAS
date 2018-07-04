{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/06/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TarifsAQui (TarifsAQui)
Mots clefs ... : TOM;TarifsAQui
*****************************************************************}
Unit TarifsAQui_TOM ;

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
  HDB,
  uTom,
  uTob,
  wTom,
  Lookup,
  Menus,
  hPanel;

Type
  TOM_yTarifsAQui = Class (twTOM)
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
    sFonctionnalite         : string;
    sAction, sDroit, sAppel : string;
    sParametreYTQ           : string;

    iDeIDentifiantYTA  : integer;
    sDeDeviseYTA       : string;

    LBNOMTIERS         : tHLabel;
    LBAFFAIRE          : tHLabel;
    LBPOLE             : tHLabel;
    LBATELIER          : tHLabel;
    LBGROUPE           : tHLabel;

    BAccesSurQuoi      : tToolbarButton97;
    BAccesConditions   : tToolbarButton97;

    TobYFO : tob;

    YTA_SITE : thEdit;
    YTA_GRP  : thEdit;

    { Click sur le menu du bouton bAccesSurQuoi }
    procedure MnSaisieGrille_OnClick(Sender: TObject);
    procedure MnSaisieOnglet_OnClick(Sender: TObject);
    procedure MnSaisieMixte_OnClick(Sender: TObject);

    { Bouton }
    procedure BACCESSURQUOI_OnClick(Sender: TOBject);
    procedure BACCESCONDITIONS_OnClick(Sender: TOBject);

    { Affichage }
    procedure SetSayOuGet(sControl: string; lConditionSay, lClef :boolean);
    procedure ControleAffichageDepotOuTarifDepot;
    procedure ControleAffichageTiersOuTarifTiers;
    procedure ControleAffichageCommercialOuTypeCommercial;
    procedure ControleAffichageAffaireOuTarifAffaire;
    procedure ControleAffichagePoleAtelierGroupe;

    procedure ClientOuFournisseur;

  protected
    { Control Field }
    procedure ControlField(FieldName: string);
    procedure CalculField(FieldName: string); 
    Function  RecordIsValid: boolean; override;
  end;

Implementation

uses
  wCommuns,
  TiersUtil,
  Tarifs,
  TarifCond,
  wJetons,
  ParamSoc
  {$IFDEF GPAO}
    ,wPilotageDeFlux
  {$ENDIF GPAO}
  ;

Const
  sFournisseur  = 'FOU'; //Nature Auxiliaire
  sClient       = 'CLI'; //Nature Auxiliaire

  sTarifFournisseur       = '101';   //Fonctionnalité : tarif fournisseur
  sCommissionFournisseur  = '102';   //Fonctionnalité : commissionnement fournisseur
  sIndirectFournisseur    = '103';   //Fonctionnalité : indirect fournisseur

  sTarifClient            = '201';   //Fonctionnalité : tarif client
  sCommissionClient       = '202';   //Fonctionnalité : commissionnement client
  sIndirectClient         = '203';   //Fonctionnalité : indirect client

  sTarifSousTraitantAchat = '301';   //Fonctionnalité : tarif sous-traitant d'achat

  sTarifSousTraitantPhase = '401';   //Fonctionnalité : tarif sous-traitant de phase

	{ Tableau des libellés de messages pour les ressources }
	TexteMessage: array[1..13] of string
          = (
	      		{ 1}  'Incohérences dans la définition de la plage de validté : Date de Début > à la Date de Fin',
            { 2}  'Conditions spéciales inconnues',
            { 3}  'Code tiers inconnu',
      			{ 4}  'Le code tiers et le code famille de tarifs tiers ne doivent pas être rensignés simultanément',
			      { 5}  'Code représentant inconnu',
			      { 6}  'Le code commercial et le code type commercial ne doivent pas être rensignés simultanément',
			      { 7}  'Cette clef naturelle existe déjà dans cette table',
            { 8}  'Code affaire inconnu',
      			{ 9}  'Le code affaire et le code famille de tarifs d''affaire ne doivent pas être rensignés simultanément',
      			{10}  'Le code dépôt et le code famille de tarifs de dépôt ne doivent pas être rensignés simultanément',
      			{11}  'Le code pôle est inconnu',
      			{12}  'Le code atelier est inconnu pour le pole',
      			{13}  'Le code groupe est inconnu pour l''atelier'
            );

	ErrDate                       =  1;
	ErrCodeTarifSpecial           =  2;
	ErrCodeTiers                  =  3;
	ErrTiersEtTarifTiers          =  4;
	ErrCodeCommercial             =  5;
	ErrCommercialEtTypeCommercial =  6;
  ErrClefNaturelle              =  7;
	ErrCodeAffaire                =  8;
	ErrAffaireEtTarifAffaire      =  9;
	ErrDepotEtTarifDepot          = 10;
  ErrCodePole                   = 11;
  ErrAtelierDuPole              = 12;
  ErrGroupeDeAtelier            = 13;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnNewRecord;
var
  TobYTA : TOB;
  TS	   : tStrings;
  iCpt   : integer;
  Q      : tQuery;
  sSql   : String;
  iAIdentifiantYTA : integer;
begin
  Inherited ;
  iAIdentifiantYTA := GetInteger('YTA_IDENTIFIANT');
  if (iDeIDentifiantYTA<>0) then
  begin
    TobYTA := Tob.Create('YTA',nil,-1);

    sSql := 'SELECT *'
          + ' FROM YTARIFSAQUI'
          + ' WHERE YTA_IDENTIFIANT = ' + IntToStr(iDeIDentifiantYTA)
          ;

    if ExisteSQL(sSql) then
    begin
      Q  := OpenSQL(sSql, True);
      TS := tStringList.Create;
      try
        TobYTA.SelectDB('',Q);
        if (TobYTA<>nil) then
        begin
          wGetTableFieldsName('YTARIFSAQUI', TS);
          for iCpt := 0 to TS.Count - 1 do
            SetField(TS[iCpt], TobYTA.GetValue(TS[iCpt]));

          SetInteger('YTA_IDENTIFIANT',iAIdentifiantYTA);
          { Trace }
          SetDateTime  ('YTA_DATECREATION', Date);
          SetDateTime  ('YTA_DATEMODIF', Date);
          SetString('YTA_CREATEUR', V_PGI.FUser);
          SetString('YTA_UTILISATEUR', V_PGI.FUser);
        end;
      finally
        Ferme(Q);
        TS.Free;
        TobYTA.Free;
      end;
    end;
  end
  else if (IKC='C') then
  begin
    SetString('YTA_FONCTIONNALITE',sFonctionnalite);
    SetString('YTA_DEVISE',V_PGI.DevisePivot);
    SetString('YTA_DATEREFERENCE','CDE');
    SetString('YTA_TTCOUHT','H');
    SetDateTime ('YTA_DATEDEBUT',iDate1900);
    SetDateTime ('YTA_DATEFIN'  ,iDate2099);
  end ;
  BAccesSurQuoi.Enabled := False;
//  Getcontrol('YTA_FONCTIONNALITE').Enabled := False;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnDeleteRecord ;
begin
  Inherited;
  wDeleteTable('yTarifsSurQuoi', 'YTQ_IDENTIFIANTYTA = "' + IntToStr(GetInteger('YTA_IDENTIFIANT')) + '"');
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnUpdateRecord ;
var
  TobYTA: Tob;
begin
  DisableControl;
  try
    if Ecran = nil then
      TobYTA := fTob
    else
    begin
    {$IFDEF EAGLCLIENT}
      TobYTA := TFFiche(Ecran).QFiche;
    {$ELSE}
      TobYTA := Tob.Create(FTableName, nil, -1);
      wMakeTobFromDS(DS, TobYTA);
    {$ENDIF}
    end;

    //Calcul du Poids du contexte
    SetInteger('YTA_POIDSRECHERCHE',CalculPoidsAQui(TobYTA));

    {$IFNDEF EAGLCIENT}
      if Assigned(Ecran) and Assigned(TobYTA) then
        TobYTA.Free;
    {$ENDIF}
  finally
    EnableControl;
  end;

  if RecordIsValid then
  begin
  end;

  Inherited ;

  BAccesSurQuoi.Enabled := True;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnAfterUpdateRecord ;
begin
  Inherited ;
  //Duplication des SurQuoi du AQui
  if (iDeIDentifiantYTA<>0) and (PgiAsk('Confirmez-vous la duplication complète du système tarifaire ?','') = mrYes) then
  begin
    DuplicationYTQ(GetString('YTA_FONCTIONNALITE'), GetString('YTA_TIERS'), GetString('YTA_TARIFTIERS'), GetString('YTA_AFFAIRE'), GetString('YTA_TARIFAFFAIRE'), iDeIDentifiantYTA, GetInteger('YTA_IDENTIFIANT'), sDeDeviseYTA, GetString('YTA_DEVISE'));
  end;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnLoadRecord ;
begin
  Inherited ;
  ControleAffichageDepotOuTarifDepot;
  ControleAffichageTiersOuTarifTiers;
  ControleAffichageCommercialOuTypeCommercial;
  ControleAffichageAffaireOuTarifAffaire;
  ControleAffichagePoleAtelierGroupe;
  SetSayOuGet('YTA_FONCTIONNALITE', True, False);
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnArgument ( S: String );
var
  iRangPanel, iHauteurDisponible : integer;

  procedure ActivePanel(sNomPanel: string; var iRangPanel: integer; var iHauteurDisponible : integer; iHauteurPanel : integer);
  begin
    //Tabsheet TSCaractérisitiques complet on active au Tabsheet TSComplements
    if (iHauteurDisponible < iHauteurPanel) then
    begin
      iHauteurDisponible := 390; //Hauteur du TabSheet
      iRangPanel         :=   6; //Premier Panel du Tabsheet complément
      tTabSheet(GetControl('TSCOMPLEMENTS')).TabVisible := True;
    end;

    thPanel(GetControl(sNomPanel)).Parent                 := tHPanel(GetControl('PN'+IntToStr(iRangPanel)));
    thPanel(GetControl('PN'+IntToStr(iRangPanel))).Height := iHauteurPanel;

    iRangPanel := iRangPanel+1;
    iHauteurDisponible := iHauteurDisponible - iHauteurPanel;
  end;

begin
  FTableName := 'YTARIFSAQUI';

  Inherited ;
  { Extraction de '[' ... ']' }
  sFonctionnalite   := GetArgumentValue(S,'YTA_FONCTIONNALITE');
  sAction           := GetArgumentValue(S,'ACTION');
  sDroit            := GetArgumentValue(S,'DROIT');

  iDeIDentifiantYTA := ValeurI(GetArgumentValue(S,'YTA_IDENTIFIANT'));
  sDeDeviseYTA      := GetArgumentValue(S,'YTA_DEVISE');

  sAppel            := GetArgumentValue(S, 'APPEL');
  sParametreYTQ     := iifV((sAppel='ARTICLE'),';APPEL='+sAppel+';YTQ_ARTICLE='+GetArgumentValue(S,'YTQ_ARTICLE')+';YTQ_TARIFARTICLE='+GetArgumentValue(S,'YTQ_TARIFARTICLE'),'');

  Ecran.Caption     := RechDom('YFONCTIONNALITES',sFonctionnalite,False)+TraduireMemoire(' - Pour Qui');
  ClientOuFournisseur;

  BAccesSurquoi      := tToolbarButton97(GetControl('BACCESSURQUOI'));
  BAccesConditions   := tToolbarButton97(GetControl('BACCESCONDITIONS'));

  LBNOMTIERS         := thLabel(GetControl('LBNOMTIERS'));
  LBAFFAIRE          := thLabel(GetControl('LBAFFAIRE'));
  LBPOLE             := thLabel(GetControl('LBPOLE'));
  LBATELIER          := thLabel(GetControl('LBATELIER'));
  LBGROUPE           := thLabel(GetControl('LBGROUPE'));

  YTA_SITE           := tHEdit(GetControl('YTA_SITE'));
  YTA_GRP            := tHEdit(GetControl('YTA_GRP') );

  if (GetControl('YTA_TARIFSPECIAL') <> nil) then SetControlProperty('YTA_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sFonctionnalite+'"');
  if (GetControl('BACCESCONDITIONS') <> nil) then TToolBarButton97(GetControl('BACCESCONDITIONS')).OnClick := BACCESCONDITIONS_OnClick ;
  if (GetControl('BACCESSURQUOI')    <> nil) then TToolBarButton97(GetControl('BACCESSURQUOI')).OnClick := BACCESSURQUOI_OnClick ;

   { Menu associé au bouton BAccesSurQuoi }
	if (GetControl('MNSAISIEGRILLE')<> nil) then TMenuItem(GetControl('MNSAISIEGRILLE')).OnClick := MnSaisieGrille_OnClick;
	if (GetControl('MNSAISIEONGLET')<> nil) then TMenuItem(GetControl('MNSAISIEONGLET')).OnClick := MnSaisieOnglet_OnClick;
	if (GetControl('MNSAISIEMIXTE') <> nil) then TMenuItem(GetControl('MNSAISIEMIXTE') ).OnClick := MnSaisieMixte_OnClick;

  SetSayOuGet('YTA_BLOCNOTE', True, False);

  BAccesConditions.Visible := GetParamSoc('SO_TARIFCONDAPPL');

  //Fixe l'écran de saisi en fonction des paramètres de gestion des informations de la fonctionnalite
  TobYFO := Tob.Create('_YFO_',nil,-1);
  GetTobYFO(sFonctionnalite, TobYFO);
  //Ces deux TabSheets regroupent toutes les données disponibles mais sont invisibles
  tTabSheet(GetControl('TSJAMAISAFFICHE1')).TabVisible := False;
  tTabSheet(GetControl('TSJAMAISAFFICHE2')).TabVisible := False;
  tTabSheet(GetControl('TSCOMPLEMENTS')).TabVisible    := False;
  //Activation du premier TabSheet de saisie
  tPageControl(GetControl('Pages')).ActivePage := tTabSheet(GetControl('PGENERAL'));
  //Sur ces TabSheets de saisis transfert des données saisissables : paramétrées comme telles dans la fonctionnalité demandée
  iRangPanel:=1; iHauteurDisponible := 390;
  if (TobYFO.Detail[0].GetValue('YFO_OKDEVISE')     =wTrue) or (TobYFO.Detail[0].GetValue('YFO_OKSPECIAL')=wTrue) then ActivePanel('PNGENERAL'          , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKDEPOT')      =wTrue)                                                       then ActivePanel('PNDEPOT'            , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKTIERS')      =wTrue)                                                       then ActivePanel('PNTIERS'            , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKAFFAIRE')    =wTrue)                                                       then ActivePanel('PNAFFAIRE'          , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKDATE')       =wTrue) or (TobYFO.Detail[0].GetValue('YFO_OKTTC')=wTrue)     then ActivePanel('PNDATE'             , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKCOMMERCIAL') =wTrue)                                                       then ActivePanel('PNCOMMERCIAL'       , iRangPanel, iHauteurDisponible,  78);
  if (TobYFO.Detail[0].GetValue('YFO_OKPOLESITEGRP')=wTrue)                                                       then ActivePanel('PNPOLEATELIERGROUPE', iRangPanel, iHauteurDisponible, 104);
  if (TobYFO.Detail[0].GetValue('YFO_OKJHL')        =wTrue)                                                       then ActivePanel('PNJOURHEURELEU'     , iRangPanel, iHauteurDisponible, 104);
  if (TobYFO.Detail[0].GetValue('YFO_OKYTALIGNE')   =wTrue)                                                       then ActivePanel('PNINFOLIGNE'        , iRangPanel, iHauteurDisponible, 104);

  thLabel(GetControl('TYTA_TARIFSPECIAL')).Visible        := (TobYFO.Detail[0].GetValue('YFO_OKSPECIAL')=wTrue);
  thDBEdit(GetControl('YTA_TARIFSPECIAL')).Visible        := (TobYFO.Detail[0].GetValue('YFO_OKSPECIAL')=wTrue);
  tWinControl(GetControl('GCASCGENERAL')).Visible         := (TobYFO.Detail[0].GetValue('YFO_OKSPECIAL') =wTrue);
  thLabel(GetControl('TYTA_DEVISE')).Visible              := (TobYFO.Detail[0].GetValue('YFO_OKDEVISE') =wTrue);
  thDBValComboBox(GetControl('YTA_DEVISE')).Visible       := (TobYFO.Detail[0].GetValue('YFO_OKDEVISE') =wTrue);
  tGroupBox(GetControl('GDATESVALIDITES')).Visible        := (TobYFO.Detail[0].GetValue('YFO_OKDATE')=wTrue);
  tWinControl(GetControl('YTA_TTCOUHT')).Visible          := (TobYFO.Detail[0].GetValue('YFO_OKTTC') =wTrue);
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnClose ;
begin
  Inherited ;
  if assigned(TobYFO) then TobYFO.Free;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.OnCancelRecord ;
begin
  Inherited ;
  BAccesSurQuoi.Enabled := True;
end ;

{--------------------------------------------------------------------------------
   Data Change Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.CalculField(FieldName: string);
var
  aTiers, aAffaire : MyArrayValue;
begin
  aTiers   := nil;
  aAffaire := nil;
//  inherited;
  //En fonction du module (nature de l'auxiliaire) affectation de la bonne liste tiers (Client ou fournisseur)
  if      (FieldName='YTA_FONCTIONNALITE') then
    ClientOuFournisseur
  else if (FieldName='YTA_TIERS') then
  begin
    if (GetString(FieldName)<>'') then
    begin
      aTiers := WGetSqlFieldsValues(['T_LIBELLE'], 'TIERS', 'T_NATUREAUXI="'+NatureAuxiliaire(GetString('YTA_FONCTIONNALITE'))+'" and T_TIERS="'+GetString(FieldName)+'"');
      if (aTiers<>nil) then
        LBNOMTIERS.Caption  := aTiers[0]
      else
        LBNOMTIERS.Caption  := TraduireMemoire('Client inconnu');
    end
    else
      LBNOMTIERS.Caption  := '';
  end
  else if (FieldName='YTA_AFFAIRE') then
  begin
    if (GetString(FieldName)<>'') then
    begin
      aAffaire := WGetSqlFieldsValues(['AFF_LIBELLE'], 'AFFAIRE', 'AFF_AFFAIRE="'+GetString(FieldName)+'"');
      if (aAffaire<>nil) then
        LBAFFAIRE.Caption  := aAffaire[0]
      else
        LBAFFAIRE.Caption  := TraduireMemoire('Affaire inconnue');
    end
    else
      LBAFFAIRE.Caption  := '';
  end
  else if (FieldName='YTA_POLE') then
  begin
    if (GetString(FieldName)<>'') then
    begin
      LBPOLE.Caption    := wGetSqlFieldValue('QPO_POLELIB'  , 'QPOLE'  , 'QPO_CTX="0" and QPO_POLE="'+GetString('YTA_POLE')+'"');
      if (LBPOLE.Caption='') then
        LBPOLE.Caption    := TraduireMemoire('Pôle inconnu');
    end
    else
      LBPOLE.Caption := '';
  end
  else if (FieldName='YTA_SITE') then
  begin
    if (GetString(FieldName)<>'') then
    begin
      LBATELIER.Caption := wGetSqlFieldValue('QSI_SITELIB'  , 'QSITE'  , 'QSI_CTX="0" and QSI_POLE="'+GetString('YTA_POLE')+'"'+' and QSI_SITE="'+GetString('YTA_SITE')+'"');
      if (LBATELIER.Caption='') then
        LBATELIER.Caption    := TraduireMemoire('Atelier inconnu');
    end
    else
      LBATELIER.Caption := '';
  end
  else if (FieldName='YTA_GRP') then
  begin
    if (GetString(FieldName)<>'') then
    begin
      LBGROUPE.Caption  := wGetSqlFieldValue('QGR_GROUPELIB', 'QGROUPE', 'QGR_CTX="0" and QGR_SITE="'+GetString('YTA_SITE')+'"'+' and QGR_GRP="' +GetString('YTA_GRP')+'"');
      if (LBGROUPE.Caption='') then
        LBGROUPE.Caption    := TraduireMemoire('Groupe inconnu');
    end
    else
      LBGROUPE.Caption := '';
  end
  ;

end;

{--------------------------------------------------------------------------------
   Validation Champ par Champ
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.ControlField(FieldName: string);
begin
	DisableControl;
  try
    if      (Pos(FieldName,'YTA_DEPOT/YTA_TARIFDEPOT')         >0) then
      ControleAffichageDepotOuTarifDepot
    else if (Pos(FieldName,'YTA_TIERS/YTA_TARIFTIERS')         >0) then
      ControleAffichageTiersOuTarifTiers
    else if (Pos(FieldName,'YTA_COMMERCIAL/YTA_TYPECOMMERCIAL')>0) then
      ControleAffichageCommercialOuTypeCommercial
    else if (Pos(FieldName,'YTA_AFFAIRE/YTA_TARIFAFFAIRE')     >0) then
      ControleAffichageAffaireOuTarifAffaire
    else if (Pos(FieldName,'YTA_POLE/YTA_SITE/YTA_GRP')        >0) then
      ControleAffichagePoleAtelierGroupe
    ;
  finally
    EnableControl;
  end;

//  inherited;

	if LastError <> 0 then
 		LastErrorMsg := TexteMessage[LastError];
end;

{--------------------------------------------------------------------------------
   Validation Global
---------------------------------------------------------------------------------}
function TOM_yTarifsAQui.RecordIsValid: boolean;
var
  sClefNaturelle : string;
  lTarifCondAppl : boolean;
begin
  // controle de la clef naturelle
  sClefNaturelle :=        '(YTA_FONCTIONNALITE="' +GetString('YTA_FONCTIONNALITE')     +'") '+
                       'and (YTA_DEPOT="'          +GetString('YTA_DEPOT')              +'") '+
                       'and (YTA_TARIFDEPOT="'     +GetString('YTA_TARIFDEPOT')         +'") '+
                       'and (YTA_DEVISE="'         +GetString('YTA_DEVISE')             +'") '+
                       'and (YTA_TARIFSPECIAL="'   +GetString('YTA_TARIFSPECIAL')       +'") '+
                       'and (YTA_COMMERCIAL="'     +GetString('YTA_COMMERCIAL')         +'") '+
                       'and (YTA_TYPECOMMERCIAL="' +GetString('YTA_TYPECOMMERCIAL')     +'") '+
                       'and (YTA_POLE="'           +GetString('YTA_POLE')               +'") '+
                       'and (YTA_SITE="'           +GetString('YTA_SITE')               +'") '+
                       'and (YTA_GRP="'            +GetString('YTA_GRP')                +'") '+
                       'and (YTA_TIERS="'          +GetString('YTA_TIERS')              +'") '+
                       'and (YTA_TARIFTIERS="'     +GetString('YTA_TARIFTIERS')         +'") '+
                       'and (YTA_AFFAIRE="'        +GetString('YTA_AFFAIRE')            +'") '+
                       'and (YTA_TARIFAFFAIRE="'   +GetString('YTA_TARIFAFFAIRE')       +'") '+
                       'and (YTA_JOURSEMAINE="'    +GetString('YTA_JOURSEMAINE')        +'") '+
                       'and (YTA_HEUREJOUR="'      +GetString('YTA_HEUREJOUR')          +'") '+
                       'and (YTA_LIEUCONSO="'      +GetString('YTA_LIEUCONSO')          +'") '+
                       'and (YTA_DATEDEBUT="'      +USDateTime(GetDateTime ('YTA_DATEDEBUT'))+'") '+
                       'and (YTA_DATEFIN="'        +USDateTime(GetDateTime ('YTA_DATEFIN'))  +'") '+
                       'and (YTA_TTCOUHT="'        +GetString('YTA_TTCOUHT')            +'") ';
  if (IKC='M') then
    sClefNaturelle := sClefNaturelle + ' and (YTA_IDENTIFIANT<>'+ IntToStr(GetInteger('YTA_IDENTIFIANT')) +')';


  lTarifCondAppl := GetParamSoc('SO_TARIFCONDAPPL');
  if (ExistYTA( sClefNaturelle )) and ((not lTarifCondAppl) or  (not (PgiAsk('Confirmez-vous l''enregistrement de cette fiche ?','ATTENTION cette clef naturelle existe déjà  !') = mrYes))) then
    LastError := ErrClefNaturelle;

  if      (GetDateTime ('YTA_DATEDEBUT')>GetDateTime ('YTA_DATEFIN'))                                                                                    then LastError := ErrDate
  else if (GetString('YTA_TARIFSPECIAL')<>'') and (not ExistTarifSpecial(GetString('YTA_FONCTIONNALITE'),GetString('YTA_TARIFSPECIAL')))       then LastError := ErrCodeTarifSpecial
  else if (GetString('YTA_TIERS')       <>'') and (not ExistTiers(NatureAuxiliaire(GetString('YTA_FONCTIONNALITE')),GetString('YTA_TIERS')))   then LastError := ErrCodeTiers
  else if (GetString('YTA_TIERS')       <>'') and (GetString('YTA_TARIFTIERS')<>'')                                                            then LastError := ErrTiersEtTarifTiers
  else if (GetString('YTA_COMMERCIAL')  <>'') and (not ExistCommercial(GetString('YTA_COMMERCIAL')))                                           then LastError := ErrCodeCommercial
  else if (GetString('YTA_COMMERCIAL')  <>'') and (GetString('YTA_TYPECOMMERCIAL')<>'')                                                        then LastError := ErrCommercialEtTypeCommercial
  else if (GetString('YTA_AFFAIRE')     <>'') and (not ExistAffaire(GetString('YTA_AFFAIRE')))                                                 then LastError := ErrCodeAffaire
  else if (GetString('YTA_AFFAIRE')     <>'') and (GetString('YTA_TARIFAFFAIRE')<>'')                                                          then LastError := ErrAffaireEtTarifAffaire
  else if (GetString('YTA_DEPOT')       <>'') and (GetString('YTA_TARIFDEPOT')<>'')                                                            then LastError := ErrDepotEtTarifDepot
 {$IFDEF GPAO}
  else if (GetString('YTA_POLE')        <>'') and (not ExistPole(GetString('YTA_POLE')))                                                       then LastError := ErrCodePole
  else if (GetString('YTA_SITE')        <>'') and (not ExistAtelierDuPole(GetString('YTA_POLE'), GetString('YTA_SITE')))                       then LastError := ErrAtelierDuPole
  else if (GetString('YTA_GRP')         <>'') and (not ExistGroupeDeAtelier(GetString('YTA_SITE'), GetString('YTA_GRP')))                      then LastError := ErrGroupeDeAtelier
  {$ENDIF}
  ;

  Result := (LastError = 0);

  if LastError > 0 then
  begin
  	LastErrorMsg := TexteMessage[LastError];
    Case LastError of
      ErrDate                       : SetFocusControl('YTA_DATEDEBUT');
      ErrCodeTarifSpecial           : SetFocusControl('YTA_TARIFSPECIAL');
      ErrCodeTiers                  : SetFocusControl('YTA_TIERS');
      ErrTiersEtTarifTiers          : SetFocusControl('YTA_TARIFTIERS');
      ErrCodeCommercial             : SetFocusControl('YTA_REPRESENTANT');
      ErrCommercialEtTypeCommercial : SetFocusControl('YTA_TYPECOMMERCIAL');
      ErrCodeAffaire                : SetFocusControl('YTA_AFFAIRE');
      ErrAffaireEtTarifAffaire      : SetFocusControl('YTA_TARIFAFFAIRE');
      ErrDepotEtTarifDepot          : SetFocusControl('YTA_TARIFDEPOT');
      ErrCodePole                   : SetFocusControl('YTA_POLE');
      ErrAtelierDuPole              : SetFocusControl('YTA_SITE');
      ErrGroupeDeAtelier            : SetFocusControl('YTA_GRP');
    end;
  end;
end;

{--------------------------------------------------------------------------------
   En fonction  de la nature  du tiers personnalisation des listes de choix et des libellés
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.ClientOuFournisseur;
var
  sNatureAuxi : string;
begin
  sNatureAuxi  := NatureAuxiliaire(sFonctionnalite);
  if      (sNatureAuxi=sFournisseur) or (sNatureAuxi=sCommissionFournisseur) then
  begin
    SetControlProperty('YTA_TIERS', 'DataType', 'GCTIERSFOURN');
    THLabel(GetControl('TYTA_TIERS')).Caption                := TraduireMemoire('Fournisseur');
    SetControlProperty('YTA_TARIFTIERS', 'DataType', 'TTTARIFFOURNISSEUR');
    THLabel(GetControl('TYTA_TARIFTIERS')).Caption           := TraduireMemoire('Famille tarif fournisseur');
  end
  else if (sNatureAuxi=sClient     ) or (sNatureAuxi=scommissionClient     ) then
  begin
    SetControlProperty('YTA_TIERS', 'DataType', 'GCTIERSCLI');
    THLabel(GetControl('TYTA_TIERS')).Caption                := TraduireMemoire('Client');
    SetControlProperty('YTA_TARIFTIERS', 'DataType', 'TTTARIFCLIENT');
    THLabel(GetControl('TYTA_TARIFTIERS')).Caption           := TraduireMemoire('Famille tarif client');
  end;
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zones Tiers et Famille Tarif Tiers
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.ControleAffichageTiersOuTarifTiers;
var
  lTiers, lNiLunNiLautre : boolean;
begin
  lTiers := (GetString('YTA_TIERS')<>'');
  lNiLunNiLautre := (not lTiers) and (GetString('YTA_TARIFTIERS')='');
  SetSayOuGet('YTA_TIERS'     , (((not lTiers) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
  SetSayOuGet('YTA_TARIFTIERS', (((    lTiers) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
//  Result := (((not lTiers) and (not lNiLunNiLautre)) or (IKC='M'));
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zones Commercial et Type Commercial
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.ControleAffichageCommercialOuTypeCommercial;
var
  lCommercial, lNiLunNiLautre : boolean;
begin
  lCommercial := (GetString('YTA_COMMERCIAL')<>'');
  lNiLunNiLautre := (not lCommercial) and (GetString('YTA_TYPECOMMERCIAL')='');
  SetSayOuGet('YTA_COMMERCIAL'    , (((not lCommercial) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
  SetSayOuGet('YTA_TYPECOMMERCIAL', (((    lCommercial) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
//  Result := (((not lCommercial) and (not lNiLunNiLautre)) or (IKC='M'));
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zones Affaire et Famille Tarif Affaire
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.ControleAffichageAffaireOuTarifAffaire;
var
  lAffaire, lNiLunNiLautre : boolean;
begin
  lAffaire := (GetString('YTA_AFFAIRE')<>'');
  lNiLunNiLautre := (not lAffaire) and (GetString('YTA_TARIFAFFAIRE')='');
  SetSayOuGet('YTA_AFFAIRE'     , (((not lAffaire) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
  SetSayOuGet('YTA_TARIFAFFAIRE', (((    lAffaire) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
//  Result := (((not lAffaire) and (not lNiLunNiLautre)) or (IKC='M'));
end;

{--------------------------------------------------------------------------------
   Controle et Affichage des zones Depot et Famille Tarif Depot
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.ControleAffichageDepotOuTarifDepot;
var
  lDepot, lNiLunNiLautre : boolean;
begin
  lDepot := (GetString('YTA_DEPOT')<>'');
  lNiLunNiLautre := (not lDepot) and (GetString('YTA_TARIFDEPOT')='');
  SetSayOuGet('YTA_DEPOT'     , (((not lDepot) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
  SetSayOuGet('YTA_TARIFDEPOT', (((    lDepot) and (not lNiLunNiLautre)) {or (IKC='M')}), False);
//  Result := (((not lDepot) and (not lNiLunNiLautre)) or (IKC='M'));
end;

{--------------------------------------------------------------------------------
   Définition des PLUS sur les Poles, les Ateliers et les groupe
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.ControleAffichagePoleAtelierGroupe;
begin
  if (YTA_SITE<>nil) and (YTA_GRP<>nil) then
  begin
    SetControlProperty('YTA_POLE', 'Plus', ''                                                    );
    SetControlProperty('YTA_SITE', 'Plus', 'QSI_CTX="0" and QSI_POLE="'+GetString('YTA_POLE')+'"');
    SetControlProperty('YTA_GRP' , 'Plus', 'QGR_CTX="0" and QGR_SITE="'+GetString('YTA_SITE')+'"');
  end;
end;

{-------------------------------------------------------------------------------
   Positionne les propriété Enable et Color d'un controle en fonction d'une condition de Say
--------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.SetSayOuGet(sControl: string; lConditionSay, lClef: boolean);
begin
  if (lConditionSay) or (sAction='CONSULTATION') or ( lClef and (DS.State<>dsInsert)) then
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
   Accès à la saisie des tarifs Sur Quoi en saisie en grille
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.MnSaisieGrille_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE=GRILLE'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Sur Quoi en saisie en onglet
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.MnSaisieOnglet_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE=ONGLET'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Sur Quoi en saisie mixte
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.MnSaisieMixte_OnClick(Sender: TObject);
begin
  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE=MIXTE'+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

{--------------------------------------------------------------------------------
   Saisie des conditions complémentaires pour l'application d'un tarif
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.BACCESCONDITIONS_OnClick(Sender: TOBject);
Var
  TheTobTarif : Tob;
  sConditionComplementaire : string;
begin
  TheTobTarif:=Tob.Create('_CONDITIONAPPLICATION_',nil,-1);
  try
    TheTobTarif.AddChampSupValeur('GF_CONDAPPLIC', GetString('YTA_BLOCNOTE'), false);
    sConditionComplementaire := GetString('YTA_BLOCNOTE');
    EntreeTarifCond(iifV(sAction='CONSULTATION',taConsult,taModif), TheTobTarif, 'Tiers');
    if (sConditioncomplementaire<>TheTobTarif.GetValue('GF_CONDAPPLIC')) then
    begin
      if (not(DS.State in [dsInsert,dsEdit])) then
        DS.Edit;
      SetString('YTA_BLOCNOTE',TheTobTarif.GetValue('GF_CONDAPPLIC'));
    end;
  finally
    TheTobTarif.Free;
  end;
end;

{--------------------------------------------------------------------------------
   Accès à la saisie des tarifs Sur Quoi selon le paramètre
---------------------------------------------------------------------------------}
procedure TOM_yTarifsAQui.BACCESSURQUOI_OnClick(Sender: TOBject);
var
  sAppelMethodeSaisie, sParamMethodeSaisie : string;
begin
  sParamMethodeSaisie := GetParamSoc('SO_TARIFSAISIE');
  if      (sParamMethodeSaisie = 'O') then sAppelMethodeSaisie := 'ONGLET'
  else if (sParamMethodeSaisie = 'G') then sAppelMethodeSaisie := 'GRILLE'
  else if (sParamMethodeSaisie = 'M') then sAppelMethodeSaisie := 'MIXTE'
  else                                     sAppelMethodeSaisie := 'MIXTE';

  AGLLanceFiche('Y','YTARIFSURQUOI_FSL',GetString('YTA_IDENTIFIANT'),'','ACTION='+DroitAction(sDroit)+';DROIT='+sDroit+';YTA_IDENTIFIANT='+GetString('YTA_IDENTIFIANT')+';METHODEDESAISIE='+sAppelMethodeSaisie+';YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_DEVISE='+GetString('YTA_DEVISE')+sParametreYTQ);
end;

Initialization
  registerclasses ( [ TOM_yTarifsAQui ] ) ;
end.
