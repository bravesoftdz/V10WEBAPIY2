{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPGENERENCADECA ()
Mots clefs ... : TOF;CPGENERENCADECA
*****************************************************************}
Unit CPGENERENCADECA_TOF ;

Interface

Uses  StdCtrls,
      Controls,
      Classes,
      Windows,
      {$IFDEF EAGLCLIENT}
        MaineAGL,
      {$ELSE}
        db,
       {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF},
        FE_Main,
        hdb,
        dbGrids,
      {$ENDIF}
      CPSaisiePiece_Tof,     // saisie paramètrable
      forms,
      sysutils,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      HTB97,                 // TToolBarButton97
      UTOF
      ,UTob
      ,Ent1            // pour ExoToDates
      ,ULibEncaDeca
      ,UObjFiltres
      ,Saisutil         // pour QuelExo, VentilerTob
      ,HQry
      ,HStatus          // pour InitMove et MoveCur et FiniMove
      ,ParamSoc         // pour le GetParamSoc
      ,CPARAMGENER_TOM  // pour CPLanceFiche_ParamGener
      ,LettUtil         // rour DateCorrecte
      ,uLibWindows      // pour TraductionTHMultiValComboBox
      ,uLibEcriture
      ,uLibPieceCompta
      ,Constantes
      ,UtilSais       // MajSoldesEcritureTOB
      , Lookup
      ;

procedure CPLanceFiche_GenereEncaDeca ( Args : string ) ;

Type
  TOF_CPGENERENCADECA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private

    FListe          : THGrid ;
    Pages           : TPageControl ;

    // Gestion des filtres
    ObjFiltre       : TObjFiltre ;

    // Controles - Onglet scénario
    FFlux           : THValComboBox ;
    FOperation      : THValComboBox ;
    FScenario       : THValComboBox ;
    FDateComptable  : THEdit ;
    FDateEche       : THEdit ;
    FAvecDateEche   : TCheckBox ;
    BParamScenario  : TButton ;

    // Controles - Onglet standard
    FTIC            : TCheckBox ;
    FCategorie      : THValComboBox ;
    FModePaiement   : THMultiValComboBox ;
    FSuiviTreso     : TCheckBox ;
    FTraiteEdite    : TCheckBox ;
    FOldKeyDown     : TKeyEvent ;

    // Variables
    Flux          : String ;          // Code du flux paramétré (DEC ou ENC)
    CodeOPE       : String ;          // Code de l'opération paramétré
    Operation     : TTypeEncaDeca ;   // Stocke les infos relatives au fonctionnement de l'opération
    TOBScenario   : TOB ;             // Données du scénario sélectionné
    FMaxLignes    : Integer ;         // Nombre max de lignes à traiter
    FListeLots    : TStringList ;     // Liste des numéros de lots à traiter
    FQteEche      : integer ;
    FBoCfonb      : Boolean ;
    FTobRefPiece  : Tob ;             // Ref des pièces crées
    FDev          : RDevise ;
    FTobCptGene   : Tob ;
    FTobLETTRAGE  : Tob ;
    FBoVersionOpti  : Boolean ;

    // EVT fiche scénario
    procedure FFluxChanged        (Sender: TObject) ;
    procedure FOperationChanged   (Sender: TObject);
    procedure FScenarioChanged    (Sender: TObject);
    procedure FAvecDateEcheClick  (Sender: TObject);

    // EVT fiche critères
    procedure MultiDossierChanged ( Sender : TObject ) ;
    procedure FTICChanged         (Sender: TObject);
    procedure E_EXERCICEChange    (Sender: TObject);
    procedure FCategorieChanged   (Sender : TObject);
    procedure FSuiviTresoChanged  (Sender : TObject);
    procedure FTraiteEditeChanged (Sender: TObject);

    // EVT fiche boutons / liste
    procedure FormKeyDown         (Sender: TObject; var Key: Word; Shift: TShiftState) ;
    procedure BParamScenarioClick (Sender: TObject);
    procedure FListeDblClick      (Sender: TObject);
    procedure BZoomClick          (Sender: TObject);

    // Initialisations
    procedure AnalyseParametre( vStArg : String ) ;
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitCriteres ;
    procedure InitScenario ;

    // Gestion scénario
    procedure ChargeOperation ;
    procedure ChargeScenario ;
    procedure majPanelScenario ;
    procedure GestionAccesScenario ;
    procedure GestionAccesDateEcheance ;
    procedure GestionAccesOperation ;
    procedure GestionAccesGeneraux ;
    procedure GestionAccesEtab ;

    // Vérification
    function  VerifieParametres : Boolean ;
    Function  DatesOk           : Boolean ;
    function  CompteEche        : Boolean ;

    // Préparation des lots
    procedure CompleteTobScenario ;
    procedure PrepareLot( lInLotEnCours : integer ) ;
//    function  ConstitueOrigine( vStLot : string ; vDossier : string = ''  ) : Tob ;
    function  ConstitueOrigineFinal( vStLot : string ; vDossier : string = ''  ) : Tob ;
//    function  TraitementLot( vStLot : string ) : Boolean ;
//    function  GenerePieceFinale( vStLot : string ) : Boolean ;

    // Traitement chargement des échéances sélectionnées
//    procedure ClickValideOld ;
    procedure ClickValideOpti ;
    procedure ClickValideOptiBis ;
    procedure ClickValide ;

    // Divers
    function  GetConditionSuiviTreso : String ;
    procedure SetChampsInvisibles ;
    function  GetNomFiltre : string ;
    function  CleEcriture( Q : TQuery ) : String ;
    function  GetConditions             : string ;
    procedure VideGrille ;
    procedure AffichePieces ;
    function  GetMaxLot : integer ;
    function  GetChampsOpti : string ;

    // Version Optimisé prélèvements DIRECT NRJ
    function  ConstitueOrigineOpti : Tob ;
    procedure SetCriteresOpti ;
    procedure InitScenarioOpti ;
    procedure InitMulOpti ;
    function  VerifieParametresOpti : Boolean ;
    procedure CPGJournalChange (Sender: TObject ) ;
    function  TraitementLotOpti( vTobOrigine : TOB ; vBoRecap : Boolean = False ) : boolean ;
    procedure RecopieEcheances( vTobEcr, vTobOrig : Tob ; vBoContrepartie : Boolean ) ;
    procedure RecupInfoCptGene ;
    procedure RenseigneLettrage( vTobEcr, vTobEch : Tob ) ;
    procedure RenseigneCFONB( vTobEcr : Tob ; vTLCfonb : Tlist ) ;
    function  EnregistrePiece( vTobGene : Tob ; vBoRecap : Boolean = False ) : boolean ;
    Procedure UpdateEcheance( vTobOrigine : Tob ) ;
    procedure InitEnteteEcr( vTobEcr : Tob ) ;
    procedure AuxiElipsisClick         ( Sender : TObject );

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  cfonb,
  UtilPGI,
  HPanel
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

//==============================================================================
procedure CPLanceFiche_GenereEncaDeca ( Args : string ) ;
begin
  AGLLanceFiche('CP', 'CPGENERENCADECA', '', '', Args ) ;
end ;
//==============================================================================

//===================================
//======== EVENEMENTS TOF ===========
//===================================

procedure TOF_CPGENERENCADECA.OnNew ;
begin
  Inherited ;
  // Placement de la 1ère page
  Pages.ActivePage        := TTabSheet(GetControl('PPARAMETRES',True)) ;
end ;

procedure TOF_CPGENERENCADECA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPGENERENCADECA.OnUpdate ;
begin
  Inherited ;
  ClickValide ;  
end ;

procedure TOF_CPGENERENCADECA.OnLoad ;
begin
  Inherited ;
  ObjFiltre.Charger ;
  SetChampsInvisibles ;
end ;

procedure TOF_CPGENERENCADECA.OnArgument (S : String ) ;
begin
  Inherited ;

  FBoVersionOpti := True ;

  // Analyse des paramètres
  AnalyseParametre( S ) ;

  // Initialisation de l'interface
  InitComposants ;
  InitEvenements ;

  // Initialisation des onglets du scénario
  InitScenario ;

  // Initialisation des critères
  InitCriteres ;

  // Pour version optimisé
  if FBoVersionOpti then
    initMulOpti ;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPGENERENCADECA.OnClose ;
begin
  Inherited ;

  if Assigned(TobScenario) then
    FreeAndNil( TOBScenario ) ;

  if Assigned(ObjFiltre)   then
    FreeAndNil(ObjFiltre);

  if Assigned(FListeLots) then
    FreeAndNil(FListeLots);

  if Assigned( FTobRefPiece ) then
    FreeAndNil( FTobRefPiece ) ;

  if Assigned( FTobCptGene ) then
    FreeAndNil( FTobCptGene ) ;

  if Assigned( FTobLETTRAGE ) then
    FreeAndNil( FTobLETTRAGE ) ;

end ;

procedure TOF_CPGENERENCADECA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPGENERENCADECA.OnCancel () ;
begin
  Inherited ;
end ;

//======================================
//======== EVENEMENTS FICHES ===========
//======================================


procedure TOF_CPGENERENCADECA.FTICChanged(Sender: TObject);
begin
  if GetControlEnabled('E_GENERAL') then //XVI 24/02/2005
     SetControlText('E_GENERAL', '') ;
  GestionAccesGeneraux ;
end;

procedure TOF_CPGENERENCADECA.FAvecDateEcheClick(Sender: TObject);
begin
  GestionAccesDateEcheance ;
end;

procedure TOF_CPGENERENCADECA.FFluxChanged(Sender: TObject);
begin
  GestionAccesOperation ;
  GestionAccesScenario ;
end;

procedure TOF_CPGENERENCADECA.FOperationChanged(Sender: TObject);
begin
  ChargeOperation ;
  GestionAccesScenario ;
end;

procedure TOF_CPGENERENCADECA.FScenarioChanged(Sender: TObject);
begin
  ChargeScenario ;
end;


procedure TOF_CPGENERENCADECA.FListeDblClick(Sender: TObject);
var AA       : TActionFiche;
    lDossier : String ;
    M        : RMVT ;
    lTob     : Tob ;
begin

  if not FListe.Enabled then Exit ;

  if FTobRefPiece.Detail.count < FListe.row then Exit ;

  lTob := FTobRefPiece.Detail[ FListe.row - 1 ] ;

  // Données d'entête
  FillChar(M,Sizeof(M),#0) ;
  M.Jal           := lTob.GetValue('E_JOURNAL') ;
  M.Exo           := lTob.GetValue('E_EXERCICE') ;
  M.DateC         := lTob.GetValue('E_DATECOMPTABLE') ;
  M.Num           := lTob.GetValue('E_NUMEROPIECE') ;
  M.Nature        := lTob.GetValue('E_NATUREPIECE') ;
  M.ModeSaisieJal := '-' ;
  M.Simul         := 'N' ;
  M.ANouveau      := False ;
  AA              := taModif ;
  if lTob.GetNumChamp('SYSDOSSIER') > 0
    then lDossier := lTob.GetValue('SYSDOSSIER')
    else lDossier := '' ;

  LanceSaisieParam( nil, AA, M, FALSE, lDossier ) ;

end;

procedure TOF_CPGENERENCADECA.BParamScenarioClick(Sender: TObject);
begin
  if V_PGI.Superviseur
    then CPLanceFiche_ParamGener('',FScenario.Value,'ACTION=MODIFICATION')
    else CPLanceFiche_ParamGener('',FScenario.Value,'ACTION=CONSULTATION') ;

  // Remise à jour des infos
  FScenario.Reload ;
  ChargeScenario ;

end;

procedure TOF_CPGENERENCADECA.E_EXERCICEChange(Sender: TObject);
begin
  if GetControlText('E_EXERCICE') = '' then
    begin
    SetControlText('E_DATECOMPTABLE',     StDate1900) ;
    SetControlText('E_DATECOMPTABLE_',    StDate2099) ;
    end
  else
    ExoToDates( GetControlText('E_EXERCICE') , THEdit(GetControl('E_DATECOMPTABLE', True)) ,
                                               THEdit(GetControl('E_DATECOMPTABLE_', True))  ) ;
end;

//====================================
//======== INITIALISATIONS ===========
//====================================

procedure TOF_CPGENERENCADECA.InitComposants;
var Composants : TControlFiltre;
begin

  // Controles - Formes
  FListe        := THGrid(        GetControl('FListe',    True) ) ;
  Pages         := TPageControl(  GetControl('Pages',     True) ) ;

  // gestion des filtres
  Composants.PopupF   := nil ;
  Composants.Filtres  := THValComboBox   (Getcontrol('FFILTRES'));
  Composants.Filtre   := TToolBarButton97(Getcontrol('BFILTRE'));
  Composants.PageCtrl := Pages ;
  ObjFiltre := TObjFiltre.Create(Composants, GetNomFiltre );

  // Controles - Onglet standard
  FTIC            := TCheckBox(     GetControl('FTIC',         True) ) ;
  FTraiteEdite    := TCheckBox(     GetControl('TraiteEdite',  True) ) ;

  // Controles - Onglet scénario
  FFlux           := THValComboBox(      GetControl('FLUX',          True) ) ;
  FOperation      := THValComboBox(      GetControl('OPERATION',     True) ) ;
  FScenario       := THValComboBox(      GetControl('SCENARIO',      True) ) ;
  FDateComptable  := THEdit(             GetControl('DATECOMPTABLE', True) ) ;
  FDateEche       := THEdit(             GetControl('DATEECHEANCE',  True) ) ;
  FAvecDateEche   := TCheckBox(          GetControl('TDATEECHEANCE', True) ) ;
  BParamScenario  := TButton(            GetControl('BPARAMSCENARIO',True) ) ;
  FCategorie      := THValComboBox(      GetControl('CATEGORIE',     True) ) ;
  FModePaiement   := THMultiValComboBox( GetControl('E_MODEPAIE',    True) ) ;
  FSuiviTreso     := TCheckBox(          GetControl('SUIVITRESO',    True) ) ;

  // Titre de la fenêtre
  Ecran.Caption := traduirememoire('Génération de masses des règlements') ;
  if Flux = 'ENC'
    then Ecran.Caption := traduirememoire('Encaissements de masse')
    else if Flux = 'DEC'
       then Ecran.Caption := Traduirememoire('Décaissements de masse') ;
  UpdateCaption(Ecran) ;

  SetControlProperty('ETABLISSEMENTPAYEUR', 'TAG', -9979 ) ;

  FListe.ListeParam:='CPZOOMMVT' ;  

end;

procedure TOF_CPGENERENCADECA.InitCriteres;
begin

  // Param onglet tables libres
  LibellesTableLibre( TTabSheet(GetControl('PCOMPLEMENT', True)), 'TT_TABLE', 'T_TABLE', 'T') ;

  // Exo et date comptable
  if VH^.CPExoRef.Code<>'' then
    begin
    SetControlText('E_EXERCICE',        VH^.CPExoRef.Code ) ;
    end
  else
    begin
    SetControlText('E_EXERCICE',        VH^.Entree.Code ) ;

    end ;

  // SBO 29/09/2004 : FQ 14692, par défaut, positionné la fourchette de date sur le mois en cours
  //                  pour éviter les temps de chargement trop long
  SetControlText('E_DATECOMPTABLE',   DateToStr( DebutDeMois(V_PGI.DateEntree) ) ) ;
  SetControlText('E_DATECOMPTABLE_',  DateToStr( FinDeMois(V_PGI.DateEntree) ) ) ;


  // Dates d'échéances
  SetControlText('E_DATEECHEANCE',    StDate1900 ) ;
  SetControlText('E_DATEECHEANCE_',   StDate2099 ) ;
  SetControlText('E_DEVISE' ,         V_PGI.DevisePivot ) ;

  // Catégorie de paiement    (par défaut <<Tous>>)
  FCategorie.ItemIndex := 0 ;
  CategorieVersModePaiement( '' , FModePaiement ) ;

  // Nature de pièce... (par défaut Avoirs et Factures)
  if Flux = 'ENC'
    then SetControlText('E_NATUREPIECE', 'AC;FC;')
    else if Flux = 'DEC'
           then SetControlText('E_NATUREPIECE', 'AF;FF;')
           else SetControlText('E_NATUREPIECE', 'AC;AF;FC;FF;');

  // Journal..(<<Tous>> par défaut)
  SetControlProperty('E_JOURNAL', 'ITEMINDEX', 0) ;

  // Suivi tréso (MIX par défaut)
  SetControlText('SUIVITRESO', 'X') ;

  GestionAccesGeneraux ;

  // On force le rib
  SetControlText('RIB','X') ;
  SetControlEnabled('RIB',False) ;

  // Cacher l'onglet contenant les champs XX_WHERE "pinvisible"
  TTabSheet(GetControl('PINVISIBLE',True)).TabVisible := not FBoVersionOpti ;

end;

procedure TOF_CPGENERENCADECA.InitEvenements;
begin

  // Ecran
  if Assigned( Ecran.OnKeyDown )
    then FOldKeyDown := Ecran.OnKeyDown
    else FOldKeyDown := nil ;
  Ecran.OnKeyDown         := FormKeyDown ;

  // liste
  FListe.OnDblClick       := FListeDblClick ;

  // Controles - Onglet scénario
  if FBoVersionOpti then
    begin
    THValComboBox(GetControl('CPGJOURNAL', True)).OnChange := CPGJournalChange ;
    end
  else
    begin
    FFlux.OnChange          := FFluxChanged ;
    FOperation.OnChange     := FOperationChanged ;
    FScenario.OnChange      := FScenarioChanged ;
    end ;
    
  FAvecDateEche.OnClick   := FAvecDateEcheClick ;
  BParamScenario.OnClick  := BParamScenarioClick ;

  // Controles - Onglet standard
  FTIC.OnClick            := FTICChanged ;
  FCategorie.OnChange     := FCategorieChanged ;
  FSuiviTreso.OnClick     := FSuiviTresoChanged ;
  ThValComboBox( GetControl('MultiDossier', True) ).OnChange := MultiDossierChanged ;
  FTraiteEdite.OnClick    := FTraiteEditeChanged ;

  // Onglet Ecritures
  THValComboBox(GetControl('E_EXERCICE',  True)).OnChange := E_EXERCICEChange ;

  // boutons...
  TButton(GetControl('BZoom', True)).OnClick := BZoomClick ;

end;

procedure TOF_CPGENERENCADECA.InitScenario;
var lStEtab : String ;
begin

  // Initialisation du Flux
  if Flux <> '' then
    begin
    FFlux.Value   := Flux ;
    FFlux.Enabled := False ;
    end
  else
    FFlux.ItemIndex       := 0 ;
  GestionAccesOperation ;

  // Initialisation de l'opération
  if CodeOPE <> '' then
    begin
    FOperation.Value    := CodeOPE ;
    FOperation.Enabled  := False ;
    end
  else
    FOperation.ItemIndex  := 0 ;
  ChargeOperation ;

  // MAj de la liste des scénario
  GestionAccesScenario ;

  // Paramètrages des dates comptable / échéance
  FDateComptable.Text   := DateToStr( V_PGI.DateEntree ) ;
  FAvecDateEche.Checked := False ;
  FDateEche.Text        := DateToStr( iDate1900 ) ;
  GestionAccesDateEcheance;

  // Initialisation de l'établissement payeur
  lStEtab := GetParamSocSecur('SO_ETABLISDEFAUT', '')  ;
  SetControlText('ETABLISSEMENTPAYEUR', lStEtab ) ;

end;


//=============================================
//======== GESTION ACCES AUC CHAMPS ===========
//=============================================

procedure TOF_CPGENERENCADECA.GestionAccesDateEcheance;
begin
  FDateEche.Enabled := FAvecDateEche.Checked ;
end;

procedure TOF_CPGENERENCADECA.GestionAccesScenario ;
begin
  if Operation.Code = 'TOUS'
    then FScenario.Plus := ' CPG_FLUXENCADECA="' + FFlux.Value + '"'
    else FScenario.Plus := ' CPG_FLUXENCADECA="' + FFlux.Value
                           + '" AND CPG_TYPEENCADECA="' + Operation.Code + '"' ;

  if GetControlText('MULTIDOSSIER')<>'' then
    FScenario.Plus := FScenario.Plus + ' AND CPG_MULTISOC="X" ' ;

  FScenario.ItemIndex := 0 ;
  ChargeScenario ;
end;



//================================
//======== CHARGEMENTS ===========
//================================

procedure TOF_CPGENERENCADECA.ChargeOperation;
begin
  if FOperation.ItemIndex = 0
    then Operation := ChargeTypeEncaDeca( 'TOUS' )
    else Operation := ChargeTypeEncaDeca( FOperation.Value ) ;
end;


procedure TOF_CPGENERENCADECA.ChargeScenario;
begin

  // Création de la TOB
  if Assigned(TobScenario) then
    FreeAndNil( TOBScenario ) ;
  TOBScenario := TOB.Create('CPARAMGENER', nil, -1) ;
  AjouteChampsSuppScenario( TOBScenario ) ;

  // Chargement des données du scenario
  if FScenario.Value <> '' then
    begin
    TOBScenario.PutValue( 'CPG_CODE' , FScenario.Value ) ;
    TOBScenario.LoadDB ;
    end ;

  // Maj descriptif dans la partie basse de l'interface
  majPanelScenario ;

  // Affichage zone de saisie de l'établissement
  GestionAccesEtab ;

end;

//===================================
//======== MAJ DES CHAMPS ===========
//===================================

procedure TOF_CPGENERENCADECA.majPanelScenario;
var lStLib : String ;
begin

  if TobScenario = nil then Exit ;

  // libellé
  SetControlText('TLIBELLE',        TOBScenario.GetValue('CPG_LIBELLE') ) ;

  // Journal et compte de génération
  lStLib:='' ;
  if trim(TOBScenario.GetValue('CPG_JOURNAL'))<>'' then
     lStLib := RechDom('TTJOURNAL', TOBScenario.GetValue('CPG_JOURNAL'), False ) ;
  SetControlText('JOURNAL',         lStLib ) ;
  SetControlText('GENERAL',         TOBScenario.GetValue('CPG_GENERAL') ) ;

  // Méthodes de génération
  SetControlText('GROUPEENCADECA',  TOBScenario.GetValue('CPG_GROUPEENCADECA') ) ;
  SetControlText('METHODEANA',      TOBScenario.GetValue('CPG_METHODEANA') ) ;

  // Avec comptabilisation
  SetControlText('COMPTABILISE',    TOBScenario.GetValue('CPG_COMPTABILISE') ) ;
  SetControlVisible('MODEPAIEMENT', TOBScenario.GetValue('CPG_COMPTABILISE')='X' ) ;
  SetControlVisible('TMODEPAIEMENT',TOBScenario.GetValue('CPG_COMPTABILISE')='X' ) ;
  if TOBScenario.GetValue('CPG_COMPTABILISE')='X' then
    begin
    // Mode de paiement
    SetControlText('MODEPAIEMENT',    TOBScenario.GetValue('CPG_MODEPAIEMENT') ) ;
    // Escompte
    SetControlText('ESCMETHODE',      TOBScenario.GetValue('CPG_ESCMETHODE') ) ;
    end ;

  // Modèle de document
  SetControlText('EDITE',           TOBScenario.GetValue('CPG_EDITE') ) ;

  //Lettrage désactivable si pièces en devise
  SetControlText('LETTRAGEAUTO',    TOBScenario.GetValue('CPG_LETTRAGEAUTO') ) ;

  // Gestion du multi-étab / multi-soc
  SetControlText('MULTIETAB',       TOBScenario.GetValue('CPG_MULTIETAB') ) ;
  SetControlText('MULTISOC',        TOBScenario.GetValue('CPG_MULTISOC') ) ;

end;


//===========================
//======== DIVERS ===========
//===========================

procedure TOF_CPGENERENCADECA.AnalyseParametre(vStArg: String);
var lStChaine : String ;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt ( vStArg ) ;

    // Flux fixé par appel
    if Pos( 'FLUX=' , Uppercase( lStChaine ) ) > 0 then
      begin
      Flux := Copy(lStChaine, 6, length( lStChaine ) ) ;
      end

    // Opération fixé par appel
    else if Pos( 'OPERATION=' , Uppercase( lStChaine ) ) > 0 then
      begin
      CodeOPE := Copy(lStChaine, 11, length( lStChaine ) ) ;
      end ;

    end ;

end;

procedure TOF_CPGENERENCADECA.BZoomClick(Sender: TObject);
begin
  FListeDblClick( nil ) ;
end;

function TOF_CPGENERENCADECA.VerifieParametres: Boolean;
begin
  Result := False ;

  // Test journal de substitution pour les pièces sans contrepartie
  if ( GetParamSocSecur( 'SO_CPJALENCADECA', '' ) = '' ) then
    begin
    PGIInfo('Vous devez renseigner le journal de substitution (paramètre sociétés - suivi-tiers) pour pouvoir utiliser la fonction.', Ecran.Caption) ;
    exit ;
    end ;

  // Test journal de substitution pour les pièces sans contrepartie
  if ( GetParamSocSecur( 'SO_CPCPTENCADECA', '' ) = '' ) then
    begin
    PGIInfo('Vous devez renseigner le compte de substitution (paramètre sociétés - suivi-tiers) pour pouvoir utiliser la fonction.', Ecran.Caption) ;
    exit ;
    end ;

  // Test scénario
  if FScenario.Value='' then
    begin
    PGIInfo('Aucun scénario sélectionné. Vous devez sélectionner un scénario.',Ecran.Caption) ;
    SetActiveTabSheet('PCRITERE') ;
    FScenario.SetFocus ;
    exit ;
    end;

  // Tests Date
  if not DatesOk then Exit ;

  // Test journal d'escompte
  if ( TobScenario.GetValue('CPG_EscMethode')<>'RIE') and ( GetParamSocSecur('SO_CPJALESCOMPTE', '') = '' ) then
    PGIInfo ( 'Le journal d''escompte n''est pas renseigné dans les paramètres société. Aucune escompte ne pourra être générée.'
             , Ecran.Caption ) ;


  if FMaxLignes = 0 then
    begin
    PGIInfo('Vous devez saisir le nombre maximum d''écheances à traiter par pièce générée.', Ecran.Caption) ;
    SetActiveTabSheet('PCRITERE') ;
    SetFocusControl('NBLIGNES') ;
    exit ;
    end ;

  if FMaxLignes > 1000 then
    begin
    if PGIAsk('Vous avez demandé le traitement de plus de 1000 échéances par pièce. Confirmez-vous le traitement ?', Ecran.Caption) <> mrYes then
      exit ;
    end ;

  // Comptage et validation
  if not CompteEche then
    exit ;

  // Tout est ok
  Result := True ;

end;




function TOF_CPGENERENCADECA.CleEcriture( Q : TQuery ) : String;
Var lStJal      : String ;
    lStExo      : String ;
    lStDate     : TDateTime ;
    lInNumP     : Integer ;
    lInNumL     : Integer ;
    lInNumEche  : Integer ;

begin

  lStJal      := Q.FindField('E_JOURNAL').AsString ;
  lStExo      := Q.FindField('E_EXERCICE').AsString ;
  lStDate     := Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  lInNumP     := Q.FindField('E_NUMEROPIECE').AsInteger ;
  lInNumL     := Q.FindField('E_NUMLIGNE').AsInteger ;
  lInNumEche  := Q.FindField('E_NUMECHE').AsInteger ;

  Result := ' WHERE E_JOURNAL="'        + lStJal                 + '" AND '
                 + 'E_EXERCICE="'       + lStExo                 + '" AND '
                 + 'E_DATECOMPTABLE="'  + UsDateTime(lStDate)    + '" AND '
                 + 'E_NUMEROPIECE='     + IntToStr(lInNumP)      + ' AND '
                 + 'E_NUMLIGNE='        + IntToStr(lInNumL)      + ' AND '
                 + 'E_NUMECHE='         + IntToStr(lInNumEche) ;

end;


procedure TOF_CPGENERENCADECA.CompleteTobScenario;
begin
  if TobScenario = nil then Exit ;

  // Récup info CFONB
  if FBoVersionOpti then
    begin
    FBoCfonb := GetControlText('EXPORTCFONB') = 'X' ;
    if FBoCfonb
      then TobScenario.PutValue   ('CPG_CFONBEXPORT', 'X' )
      else TobScenario.PutValue   ('CPG_CFONBEXPORT', '-' ) ;
    end
  else
    FBoCfonb := TobScenario.GetValue('CPG_CFONBEXPORT') = 'X' ;
  if FBoCfonb then
    TobScenario.PutValue   ('CFONBBATCH', 'X' ) ;

  // MAJ indicateurs zappés par la gestion de masse
  TobScenario.putValue('CPG_EDITE',          '-') ;
  TobScenario.putValue('CPG_BORDEREAUEXP',   '-') ;
  TobScenario.putValue('CPG_AVECVERIFPIECE', '-') ;

  // Paramètres contextuelles
  TobScenario.PutValue('DATECOMPTABLE',   StrToDate( FDateComptable.Text ) ) ;
  if FAvecDateEche.Checked
    then TobScenario.PutValue('FORCERDATEECHE', 'X' )
    else TobScenario.PutValue('FORCERDATEECHE', '-' ) ;
  TobScenario.PutValue('DATEECHE',        StrToDate( FDateEche.Text ) ) ;
  TobScenario.PutValue('ETABLISSEMENTPAYEUR',  GetControlText('ETABLISSEMENTPAYEUR') ) ;

  // Paramètres d'éditions
  TobScenario.PutValue('APERCU',               '-' ) ;
  TobScenario.PutValue('SPOOLER',              '-' ) ;
  TobScenario.PutValue('XFICHIERSPOOLER',      '-' ) ;
  TobScenario.PutValue('REPSPOOLER',           '' ) ;

  // Paramètres d'export
  TobScenario.PutValue('EXPORTFICHIER',        '' ) ;
  TobScenario.PutValue('EXPORTAPERCU',         '-' ) ;
  TobScenario.PutValue('REMISEREF',            '' ) ;

  // Sélection uniquement de Tic/Tid
  if GetCheckBoxState('FTIC') = cbChecked
    then TobScenario.PutValue('TICTID', 'X' )
    else TobScenario.PutValue('TICTID', '-' ) ;

  // Nb d'échéance max à traiter par pièces
  if FMaxLignes > 0
    then TobScenario.PutValue('NBLIGNES', FMaxLignes )
    else TobScenario.PutValue('NBLIGNES', 0 ) ;

  TobScenario.PutValue   ('ENCADECABATCH',    'X' ) ;

end;

procedure TOF_CPGENERENCADECA.FCategorieChanged(Sender : TObject);
begin
  //  remplissage du combo des mode de paiement
  if FCategorie.ItemIndex = 0
     then CategorieVersModePaiement( '' , FModePaiement )
     else CategorieVersModePaiement( FCategorie.Value , FModePaiement ) ;
end;



procedure TOF_CPGENERENCADECA.GestionAccesOperation;
Begin
  //on ne prend que les type d'enca/Déca que corresponent....
  SetControlProperty('OPERATION','PLUS',ConditionTypeEncaDeca(FFlux.Value)) ;
  if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
  Begin
    FOperation.ItemIndex:=0 ;
    GestionAccesScenario ;
  End ;
end;


procedure TOF_CPGENERENCADECA.GestionAccesGeneraux;
begin
  SetControlEnabled('E_AUXILIAIRE',  not FTIC.Checked ) ;
  SetControlEnabled('TE_AUXILIAIRE', not FTIC.Checked ) ;
  SetControlEnabled('SOLDEPOSITIF',  not FTIC.Checked ) ;
  if FTIC.Checked then
    begin
    SetControlText('E_AUXILIAIRE', '' ) ;
    SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGTIDTIC') ;
    // restriction sur solde du tiers
    SetControlText('SOLDEPOSITIF', '-' ) ;
    end
  else SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGCOLLECTIF') ;
  SetControlProperty('E_GENERAL', 'PLUS', ' AND ' + GetConditionSuiviTreso ) ;
end;

procedure TOF_CPGENERENCADECA.FSuiviTresoChanged(Sender: TObject);
begin
  if GetControlEnabled('E_GENERAL') then
    SetControlText('E_GENERAL', '') ;
  GestionAccesGeneraux ;
end;


function TOF_CPGENERENCADECA.GetConditionSuiviTreso: String;
begin
  if GetControlText('SUIVITRESO')='X'
    then Result := 'G_SUIVITRESO IN ("' + GetControlText('FLUX')  + '","MIX")'
    else Result := 'G_SUIVITRESO = "' + GetControlText('FLUX') + '"' ;
end;

procedure TOF_CPGENERENCADECA.MultiDossierChanged(Sender: TObject);
begin
  GestionAccesScenario ;
end;


procedure TOF_CPGENERENCADECA.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

{  if ( csDestroying in Ecran.ComponentState ) then Exit ;

  case Key of

   VK_F10 : begin
            ClickValide;
            Key := 0;
            end;

    end; // case
}
  if Assigned( FOldKeyDown ) then
     FOldKeyDown(Sender, Key, Shift);

end;

procedure TOF_CPGENERENCADECA.GestionAccesEtab;
var lBoAvecEtab : Boolean ;
begin
  lBoAvecEtab := ( TobScenario <> nil ) and ( TobScenario.GetValue('CPG_MULTIETAB')='X' )
                                        and ( TobScenario.GetValue('CPG_SELECTETAB')='FAC' ) ;
  SetControlVisible('TETABLISSEMENTPAYEUR', lBoAvecEtab ) ;
  SetControlVisible('ETABLISSEMENTPAYEUR', lBoAvecEtab ) ;
end;

function TOF_CPGENERENCADECA.DatesOk: Boolean;
Var DD  : TDateTime ;
    Err : integer ;
begin
  Result := FALSE ;

  if Not IsValidDate( FDateComptable.Text ) then
    begin
    PGIError('Vous devez renseigner une date valide.', Ecran.caption ) ;
    SetActiveTabSheet('PCRITERE') ;
    FDateComptable.Text := DateToStr(V_PGI.DateEntree) ;
    FDateComptable.SetFocus ;
    Exit ;
    end
  else
    begin
    // FQ 17165 : contrôle fourchette de dates
    Err := ControleDate( FDateComptable.Text ) ;
    if Err>0 then
      begin
      Case Err of
        1 : PGIError('La date que vous avez renseignée n''est pas valide.', Ecran.caption ) ;
        2 : PGIError('La date que vous avez renseignée n''est pas dans un exercice ouvert.', Ecran.caption ) ;
        3,4 : PGIError('La date que vous avez renseignée est antérieure à une clôture.', Ecran.caption ) ;
        5 : PGIError('La date que vous avez renseignée est en dehors des limites autorisées.', Ecran.caption ) ;
        else
        end ;
      SetActiveTabSheet('PCRITERE') ;
      FDateComptable.Text := DateToStr(V_PGI.DateEntree) ;
      FDateComptable.SetFocus ;
      Exit ;
      end
    else
      begin
      DD  := StrToDate(FDateComptable.Text) ;
      if RevisionActive(DD) then
        begin
        SetActiveTabSheet('PCRITERE') ;
        FDateComptable.Text := DateToStr(V_PGI.DateEntree) ;
        FDateComptable.SetFocus ;
        Exit ;
        end ;
      // FQ 17165 : contrôle fourchette de dates
      if FAvecDateEche.Checked then
        if Not IsValidDate( FDateEche.Text ) then
          begin
          PGIError('Vous devez renseigner une date valide.', Ecran.caption ) ;
          SetActiveTabSheet('PCRITERE') ;
          FDateEche.SetFocus ;
          Exit;
          end
        else
          if not NbJoursOk( DD, StrToDate( FDateEche.Text ) ) then
            begin
            PGIError('La date d''échéance doit respecter la plage de saisie autorisée', Ecran.Caption) ;
            SetActiveTabSheet('PCRITERE') ;
            FDateEche.SetFocus ;
            Exit;
            end ;
      end ;
    end ;

  Result := TRUE ;

end;


procedure TOF_CPGENERENCADECA.SetChampsInvisibles;
var lStWhere   : String ;
    lStCond    : String ;
    lStBidon   : String ; // chaine bidon pour passage arguments var
begin

  lStWhere := '' ;
  lStBidon := '' ;
  lStCond  := '' ;

  // ===================================================
  // MAJ des champs invisibles en fonctions des critères
  // ===================================================

  // TIC / TID
  if GetControlText('FTIC') = 'X'
     then SetControlText('XX_WHEREAUXI', ' G_NATUREGENE IN ("TIC","TID") ')
     else SetControlText('XX_WHEREAUXI', ' E_AUXILIAIRE<>"" ') ;

  // Condition "Traite Editée"
  Case FTraiteEdite.State of
    cbChecked :   lStWhere := ' E_NUMTRAITECHQ<>"" ' ;
    cbUnchecked : lStWhere := ' E_NUMTRAITECHQ="" ' ;
  end ;

  // RIB
  if GetControlText('RIB') = 'X'
     then lStWhere := lStWhere + ' AND E_RIB<>"" ' ;


  // condition Sur mode de paiement : Si tous les modes d'une catégorie précise, ajouter filtre sur la liste des modes
  if (FCategorie.ItemIndex > 0) and FModePaiement.Tous and ( FModePaiement.Values.Count > 0 ) then
    begin
    TraductionTHMultiValComboBox( FModePaiement, lStCond, lStBidon, 'E_MODEPAIE', False ) ;
    lStWhere := lStWhere + ' AND ' + lStCond ;
    end ;

  // condition sur le Suivi de trésorerie
  lStWhere := lStWhere + ' AND ' + GetConditionSuiviTreso ;

  // MAJ XX_WHERE
  SetControlText('XX_WHERE', lStWhere ) ;

end;

function TOF_CPGENERENCADECA.GetNomFiltre: string;
begin
  if Flux <> ''
    then result := 'CPGENER' + Flux
    else result := 'CPGENERENCADECA' ;
end;
(*
procedure TOF_CPGENERENCADECA.ClickValideOld;
var i              : integer ;
    lStLot         : string ;
    lStJal         : string ;
    lStCpt         : string ;
    lStFin         : string ;
    lboTransacOk   : Boolean ;
begin

  FMaxLignes    := ValeurI(GetControlText('NBLIGNES'));
  lStJal        := TobScenario.Getstring('CPG_JOURNAL' ) ;
  lStCpt        := TobScenario.Getstring('CPG_GENERAL' ) ;
  lboTransacOk  := False ;

  // Maj des champs supplémentaires de paramètrage
  CompleteTobScenario ;

  // Test d'usage
  if not VerifieParametres then Exit ;

  // Vide Grille
  if Assigned(FTobRefPiece)
    then FTobRefPiece.ClearDetail
    else FTobRefPiece  := Tob.Create('$REFPIECE', nil, -1 ) ;
  VideGrille ;

  // Recup info CFONB
  if FBoCfonb then
    if not ParamCFONB( TobScenario ) then Exit ;
{
  try
    try
}
      // Préparation des lots
      PrepareLot( 1 ) ;
      if FListeLots.count = 0 then
        begin
        PgiInfo('Aucun lot n''a été généré ! Veuillez revoir votre sélection.', Ecran.Caption ) ;
        exit ;
        end
      // (plus d'un lot...)
      else if (FListeLots.count > 1) and (TobScenario.GetString('CPG_GROUPEENCADECA') = 'GLO') then
        begin
        // Si génération globalise, positionnement sur jal et cpt de substitution
        TobScenario.PutValue('CPG_JOURNAL', GetParamSocSecur( 'SO_CPJALENCADECA', '') ) ;
        TobScenario.PutValue('CPG_GENERAL', GetParamSocSecur( 'SO_CPCPTENCADECA', '') ) ;
        end ;

      // Exécution du traitement sur les lots
      For i := 0 to FListeLots.count - 1 do
        begin

        // quel lot à traiter ?
        if FBoCfonb then
          begin
          if i = (FListeLots.count - 1)
            then TobScenario.PutValue('CFONB_DERNIERLOT', 'X' )
            else if i > 0
              then TobScenario.PutValue('CFONB_PREMIERLOT', '-' ) ;
          end ;

        SetControlText('MESSAGE', 'Génération du lot n° ' + IntToStr( i+1 ) + ' / ' + IntToStr ( FListeLots.Count ) + '...' ) ;
        lStLot := FListeLots[i] ;
        lBoTransacOk := TraitementLot( lStLot ) ;

        end ;

      // Traitement pièce finale
      if (FListeLots.count > 1) and (TobScenario.GetString('CPG_GROUPEENCADECA') = 'GLO') then
        begin
        TobScenario.PutValue('CPG_JOURNAL', lStJal ) ;
        TobScenario.PutValue('CPG_GENERAL', lStCpt ) ;
        SetControlText('MESSAGE', 'Génération de la pièce finale...' ) ;
        lBoTransacOk := GenerePieceFinale( 'RECAP ' + V_PGI.User )  ;
        end ;
{
      except
        on E: Exception do
          PgiError('Erreur lors du traitement : ' + E.Message, Ecran.caption ) ;
      end ;  // fin try finally

    finally
}
    // MAJ du champ lot
    if FListeLots.count > 0 then
      begin
      SetControlText('MESSAGE', 'Maj des références des lots...' ) ;
      lStFin := TobScenario.GetValue('CPG_FLUXENCADECA') + '-' + V_PGI.User + '-' + FormatDateTime('yyyymmdd',Date)  ;
      lStFin := Copy( lStFin, 1, 17 ) ;
      For i := 0 to FListeLots.count - 1 do
        begin
        lStLot := FListeLots[i] ;
        ExecuteSQL('UPDATE ECRITURE SET E_NOMLOT="' + lStFin + '" WHERE E_NOMLOT="' + lStLot + '"' ) ;
        end ;
      end ;
//   end ; // fin try finally

  // Messages
  SetControlText('MESSAGE', 'Liste des pièces générées' ) ;

  if FBoCfonb then
    begin

    if (TobScenario.GetNumChamp('CFONB_ERREUR') > 0) and (TobScenario.GetValue('CFONB_ERREUR') > 0 )  then
      begin
      if TobScenario.GetValue('CFONB_ERREUR') = 16 then
        PgiInfo('Certains lots n''ont pas pu être exportés !', Ecran.Caption) ;
      end

    else if (TobScenario.GetNumChamp('CFONB_EXISTEMAUVAIS') > 0) and (TobScenario.GetValue('CFONB_EXISTEMAUVAIS') = 'X' )  then
        PgiInfo('Vous avez des RIB incomplets et/ou des montants de mauvais sens. Ces éléments ont été rejetés.')

    else if lBoTransacOk then
      PGIInfo('Les pièces et le ficher CFONB ont été correctement générés.', Ecran.Caption) ;
    end
  else if lBoTransacOk
    then PGIInfo('Le traitement n''a pas pu aboutir.', Ecran.Caption)
    else PGIInfo('Les pièces comptables ont été correctement générées.', Ecran.Caption) ;

end;
*)
procedure TOF_CPGENERENCADECA.FTraiteEditeChanged(Sender: TObject);
begin
  SetControlEnabled('E_NUMTRAITECHQ', FTraiteEdite.State = cbChecked ) ;
  SetControlEnabled('E_NUMTRAITECHQ_', FTraiteEdite.State = cbChecked ) ;
  SetControlEnabled('TE_NUMTRAITECHQ_', FTraiteEdite.State = cbChecked ) ;
end;

function TOF_CPGENERENCADECA.CompteEche: Boolean;
var lQEche : TQuery ;
    lSolde : Double ;
begin

  result   := False ;
  FQteEche := 0 ;

  lQEche := openSQL( 'SELECT COUNT(*) QTE, Sum(E_DEBIT) DEB, SUM(E_CREDIT) CRE FROM SUIVIMP '
                     + GetConditions , True ) ;
  if (lQEche.Eof) or (lQEche.FindField('QTE').AsInteger = 0) then
    begin
    PGIInfo('Le résultat de la recherche est vide. Veuillez revoir vos critères.', Ecran.Caption) ;
    end
  else
    begin
    if Flux = 'ENC'
      then lSolde := Arrondi(lQEche.FindField('DEB').AsFloat - lQEche.FindField('CRE').AsFloat, V_PGI.OkDecV)
      else lSolde := Arrondi(lQEche.FindField('CRE').AsFloat - lQEche.FindField('DEB').AsFloat, V_PGI.OkDecV) ;
    FQteEche := lQEche.FindField('QTE').AsInteger ;
    result := PgiAskCancel( 'Votre sélection comporte ' + IntToStr(FQteEche) + ' échéances pour un solde total de '
                                                        + StrFMontant( lSolde, 15, V_PGI.OkDecV, '', True )
                            + '. Voulez-vous appliquer le scénario ?' , Ecran.Caption ) = mrYes;
    end ;
  Ferme( lQEche ) ;

end;

function TOF_CPGENERENCADECA.GetConditions: string;
begin
  result := RecupWhereCritere( Pages ) ;
end;

procedure TOF_CPGENERENCADECA.PrepareLot( lInLotEnCours : integer ) ;
var lQEcr     : TQuery ;
    lStNumLot : string ;
    lInQte    : integer ;
begin

  if lInLotEnCours = 1 then
    begin
    if not Assigned( FListeLots ) then
      FListeLots := TStringList.Create ;
    FListeLots.Clear ;
    InitMove( GetMaxLot, ''  ) ;
    end ;

  lQEcr := OpenSQL( 'SELECT ##TOP ' + IntToStr( FMaxLignes ) + '## E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE '
                    + ' FROM SUIVIMP ' + GetConditions + ' AND E_NOMLOT=""', True ) ;

  // On ne récupère plus rien, on a atteind la fin de la sélection
  if lQEcr.Eof then
    begin
    FiniMove ;
    Ferme(lQEcr) ;
    Exit ;
    end ;

  SetControlText('MESSAGE', 'Préparation du lot n° ' + IntToStr(lInLotEnCours) + ' / ' + IntToStr ( GetMaxLot ) + '...' ) ;

  // Calcul Code Lot
  lStNumLot := V_PGI.User + Copy('00000', 1, 5 - length( IntToStr( lInLotEnCours ) ) ) + IntToStr( lInLotEnCours ) ;
  lInQte    := 0 ;

  // Traitement
  while not lQEcr.Eof do
    begin
    ExecuteSQL('UPDATE ECRITURE SET E_NOMLOT = "' + lStNumLot + '" ' + CleEcriture(lQEcr) + ' AND E_NOMLOT="" ' ) ;
    lQEcr.next ;
    Inc(lInQte) ;
    end ;
  Ferme( lQEcr ) ;

  // Stockage Code Lot
  FListeLots.Add( lStNumLot ) ;

  // Suivante
  MoveCur(False) ;

  if lInQte >= FMaxLignes then
    PrepareLot( lInLotEnCours + 1 ) ;

end;
(*
function TOF_CPGENERENCADECA.ConstitueOrigine(vStLot: string ; vDossier : string = '' ): Tob;
Var //lQEcr  : TQuery ;
    lStReq : String ;
begin
  Result := TOB.Create( 'TOB_ORIGINE', nil, -1 ) ;
{
  lStReq := 'SELECT ECRITURE.*, T_NATUREAUXI NATUREAUXI, G_NATUREGENE NATUREGENE'
               + ' FROM ECRITURE'
                   + ' LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE'
                   + ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL '
               + ' WHERE E_NOMLOT="' + vStLot + '"' ;
}
  lStReq := 'SELECT ' + GetChampsOpti
            + ' FROM ' + GetTableDossier( vDossier, 'ECRITURE')
            + ' WHERE E_NOMLOT="' + vStLot + '"' ;

  result.LoadDetailFromSQL( lStReq );
{
  lQEcr  := OpenSelect( lStReq , vDossier ) ;
  if not lQEcr.Eof then
//    result.LoadDetailDB( 'ECRITURE', '', '', lQEcr, True, True ) ;
  Ferme(lQEcr) ;
}
end;
*)
(*
function TOF_CPGENERENCADECA.TraitementLot( vStLot : string ) : Boolean ;
var lTOBOrigine    : TOB ;             // Données sur les pièces d'origine
begin

  try

    // Constitution TobOrigine
    lTOBOrigine := ConstitueOrigine( vStLot ) ;

    // indicateur scenario
    TobScenario.PutValue      ('NOMLOT',  vStLot  ) ;

    // Traitement TobOrigine
    result := ExecuteEncaDeca( lTobOrigine , TobScenario, FTobRefPiece ) ;

    // Affichage de la pièce générée
    if result then
      AffichePieces ;

    except
      result := False ;
    end ;

  if Assigned(lTOBOrigine) then
    FreeAndNil(lTOBOrigine) ;

end;
 *)
procedure TOF_CPGENERENCADECA.VideGrille;
begin
  FTobRefPiece.PutGridDetail( FListe, True, True, 'E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_DEBIT;E_CREDIT', True ) ;
end;

procedure TOF_CPGENERENCADECA.AffichePieces;
begin
  FTobRefPiece.PutGridDetail( FListe, True, True, 'E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_DEBIT;E_CREDIT', True ) ;
end;
(*
function TOF_CPGENERENCADECA.GenerePieceFinale( vStLot : string ) : Boolean ;
var lTOBOrigine    : TOB ;             // Données sur les pièces d'origine
begin

  try

    // Constitution TobOrigine
    lTOBOrigine := ConstitueOrigineFinal( vStLot ) ;

    // indicateur scenario
    TobScenario.PutValue      ('NOMLOT',  '' ) ;
    TobScenario.PutValue      ('CPG_CFONBEXPORT', '-' ) ;

    // Doit-on provoquer le lettrage ?
    TobScenario.PutValue('SANSLETTRAGE', 'X') ;
    {
    if ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + GetParamSocSecur('SO_CPCPTENCADECA', '') + '" AND G_LETTRABLE="X"')
      then TobScenario.PutValue('SANSLETTRAGE', '-')
      else TobScenario.PutValue('SANSLETTRAGE', 'X') ;
    }
    // Traitement TobOrigine
    result := ExecuteEncaDeca( lTobOrigine , TobScenario, FTobRefPiece ) ;

    // Affichage de la pièce générée
    if result then
      AffichePieces ;

    except
      result := False ;
  end ;

  // suite
  if Assigned( lTOBOrigine ) then
    FreeAndNil(lTOBOrigine) ;

end;
*)
function TOF_CPGENERENCADECA.ConstitueOrigineFinal( vStLot : string ; vDossier: string ) : Tob ;
Var //lQEcr  : TQuery ;
    lStReq : String ;
//    lStJal : String ;
//    lStCpt : String ;
begin
  Result := TOB.Create( 'TOB_ORIGINE', nil, -1 ) ;

  {
  lStCpt := GetParamSocSecur('SO_CPCPTENCADECA', '') ;
  lStJal := GetParamSocSecur('SO_CPJALENCADECA', '') ;
  lStReq := 'SELECT ECRITURE.*, T_NATUREAUXI NATUREAUXI, G_NATUREGENE NATUREGENE'
               + ' FROM ECRITURE'
                   + ' LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE'
                   + ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL '
               + ' WHERE E_GENERAL="' + lStCpt + '" AND E_JOURNAL="' + lStJal + '"'
                 + ' AND E_NOMLOT LIKE "RECAP ' + V_PGI.User + '"' ;

}
  lStReq := 'SELECT ' + GetChampsOpti + ', ' + TPieceCompta.GetChampsAuxi + ', ' + TPieceCompta.GetChampsGene
               + ' FROM ECRITURE'
                   + ' LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE'
                   + ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL '
            + ' WHERE E_NOMLOT="' + vStLot + '"' ;

  result.LoadDetailFromSQL( lStReq ) ;

{
  lQEcr  := OpenSelect( lStReq , vDossier ) ;
  if not lQEcr.Eof then
    begin
    result.LoadDetailDB( 'ECRITURE', '', '', lQEcr, True, True ) ;
    FListeLots.Add('RECAP ' + V_PGI.User ) ;
    end ;
  Ferme(lQEcr) ;
}

end;

function TOF_CPGENERENCADECA.GetMaxLot: integer;
begin
  if FQteEche <= FMaxLignes
    then result := 1
    else if (FQteEche mod FMaxLignes) = 0
     then result := (FQteEche div FMaxLignes)
     else result := (FQteEche div FMaxLignes) + 1 ;
end;

function TOF_CPGENERENCADECA.GetChampsOpti: string;
begin
  result :=  'E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_GENERAL, E_AUXILIAIRE, E_DEBIT, E_CREDIT, '
           + 'E_DEBITDEV, E_CREDITDEV, E_COUVERTURE, E_COUVERTUREDEV, E_ECHE, E_NUMECHE, E_ANA, E_EXERCICE, '
           + 'E_NOMLOT, E_NATUREPIECE, E_NUMLIGNE, E_NUMTRAITECHQ, E_ETABLISSEMENT, E_MODEPAIE, '
           + 'E_REFINTERNE, E_LIBELLE, E_RIB, E_TVAENCAISSEMENT, E_REGIMETVA, E_TVA, E_TPF, E_DATEECHEANCE, '
           + 'E_REFEXTERNE, E_DATEREFEXTERNE, E_REFLIBRE, E_AFFAIRE, E_QTE1, E_QTE2, E_QUALIFQTE1, '
           + 'E_QUALIFQTE2, E_LIBRETEXTE0, E_LIBRETEXTE1, E_LIBRETEXTE2, E_LIBRETEXTE3, E_LIBRETEXTE4, '
           + 'E_LIBRETEXTE5, E_LIBRETEXTE6, E_LIBRETEXTE7, E_LIBRETEXTE8, E_LIBRETEXTE9, E_TABLE0, '
           + 'E_TABLE1, E_TABLE2, E_TABLE3, E_LIBREDATE, E_LIBREBOOL0, E_LIBREBOOL1, E_LIBREMONTANT0, '
           + 'E_LIBREMONTANT1, E_LIBREMONTANT2, E_LIBREMONTANT3, E_CONSO, E_DEVISE, E_ORIGINEPAIEMENT, '
           + 'E_ENCAISSEMENT, E_SAISIMP, E_CFONBOK, E_QUALIFPIECE, E_ECRANOUVEAU, E_NUMENCADECA, '
           + 'E_CONTREPARTIEGEN, E_CONTREPARTIEAUX, E_BANQUEPREVI, E_QUALIFORIGINE, E_TAUXDEV, '
           + 'E_EDITEETATTVA, E_LETTRAGE, E_ETATLETTRAGE, E_DATEPAQUETMIN, E_DATEPAQUETMAX, E_TRESOSYNCHRO ' ;
end;

procedure TOF_CPGENERENCADECA.ClickValideOpti;
var lInLot         : integer ;
    lStJal         : string ;
    lStCpt         : string ;
    lboTransacOk   : Boolean ;
    lTobOrigine    : TOB ;
begin

  // MAJ des critères cachés
  SetCriteresOpti;

  FMaxLignes    := ValeurI(GetControlText('NBLIGNES'));
  lboTransacOk  := False ;

  // Maj des champs supplémentaires de paramètrage
  InitScenarioOpti ;
  lStJal        := TobScenario.Getstring('CPG_JOURNAL' ) ;
  lStCpt        := TobScenario.Getstring('CPG_GENERAL' ) ;
  CompleteTobScenario ;
  CompleteScenario ( TobScenario ) ;

  // recup info devise
  FDev.Code  := GetControlText('E_DEVISE') ;
  if FDev.Code='' then
    FDev.Code := V_PGI.DevisePivot ;
  GetInfosDevise( FDev ) ;
  FDev.Taux := GetTaux( FDev.Code, FDev.DateTaux, TobScenario.GetValue('DATECOMPTABLE') ) ;

  // Test d'usage
  if not VerifieParametresOpti then Exit ;

  // Vide Grille
  if Assigned(FTobRefPiece)
    then FTobRefPiece.ClearDetail
    else FTobRefPiece  := Tob.Create('$REFPIECE', nil, -1 ) ;
  VideGrille ;

  // Recup info CFONB
  if FBoCfonb then
    if not ParamCFONB( TobScenario ) then Exit ;

  try

    lInLot := 0 ;

    repeat

     Inc( lInLot ) ;
     // indicateurs CFONB
     if FBoCfonb and (lInLot > 1) then
       TobScenario.PutValue('CFONB_PREMIERLOT', '-' ) ;
     // Message
     SetControlText('MESSAGE', 'Génération du lot n° ' + IntToStr( lInLot ) + ' / ' + IntToStr ( GetMaxLot ) + '...' ) ;
     Application.ProcessMessages ;

     // Récup des échéances d'origine
     if Assigned( lTobOrigine ) then
       begin
       lTobOrigine.ClearDetail ;
       FreeAndNil( lTobOrigine ) ;
       end ;
     lTobOrigine := ConstitueOrigineOpti ;

     // traitement des lots :
     if lTobOrigine.detail.count > 0 then
       begin

       // Si génération globalise multi-lot, positionnement sur jal et cpt de substitution
       if lTobOrigine.detail.count >= FMaxLignes then
         begin
         if lInLot = 1 then
           begin
           TobScenario.PutValue('CPG_JOURNAL', GetParamSocSecur( 'SO_CPJALENCADECA', '') ) ;
           TobScenario.PutValue('CPG_GENERAL', GetParamSocSecur( 'SO_CPCPTENCADECA', '') ) ;
           TobScenario.PutValue('TRESOSYNCHRO', ets_Rien ) ;
           end ;
         end
       else
         begin
         if lInLot = 1 then
           TobScenario.PutValue( 'NOMLOT',  ''  ) ;
         if FBoCfonb then
           TobScenario.PutValue('CFONB_DERNIERLOT', 'X' ) ; // indicateurs CFONB
         end ;

       // Traitement lot
       lBoTransacOk := TraitementLotOpti( lTobOrigine ) ;
       end ;

     until ( lTobOrigine.Detail.count = 0 ) ;

     // Traitement pièce finale
     if (lInLot > 1) and (TobScenario.GetValue( 'NOMLOT' ) <> '') then
       begin
       TobScenario.PutValue('CPG_JOURNAL', lStJal ) ;
       TobScenario.PutValue('CPG_GENERAL', lStCpt ) ;
       TobScenario.PutValue('CPG_CFONBEXPORT', '-' ) ;
       TobScenario.PutValue('TRESOSYNCHRO', ets_Nouveau ) ;

       SetControlText('MESSAGE', 'Génération de la pièce finale...' ) ;
       Application.ProcessMessages ;

       if Assigned( lTobOrigine ) then
         begin
         lTobOrigine.ClearDetail ;
         FreeAndNil( lTobOrigine ) ;
         end ;
       lTOBOrigine  := ConstitueOrigineFinal( TobScenario.GetValue( 'NOMLOT' ) ) ;
       lBoTransacOk := TraitementLotOpti( lTobOrigine, True ) ;

      // Effacement du nom de lot temporaire
      ExecuteSQL('UPDATE ECRITURE SET E_NOMLOT="" WHERE E_NOMLOT="' + TobScenario.GetValue( 'NOMLOT' ) + '"' ) ;

       end ;


    except
      on E: Exception do
        PgiError('Erreur lors du traitement : ' + E.Message, Ecran.caption ) ;
    end ;

  if Assigned( lTobOrigine ) then
     begin
     lTobOrigine.ClearDetail ;
     FreeAndNil( lTobOrigine ) ;
     end ;

  // Messages
  SetControlText('MESSAGE', 'Liste des pièces générées' ) ;

  if FBoCfonb then
    begin

    if (TobScenario.GetNumChamp('CFONB_ERREUR') > 0) and (TobScenario.GetValue('CFONB_ERREUR') > 0 )  then
      begin
      if TobScenario.GetValue('CFONB_ERREUR') = 16 then
        PgiInfo('Certains lots n''ont pas pu être exportés !', Ecran.Caption) ;
      end

    else if (TobScenario.GetNumChamp('CFONB_EXISTEMAUVAIS') > 0) and (TobScenario.GetValue('CFONB_EXISTEMAUVAIS') = 'X' )  then
        PgiInfo('Vous avez des RIB incomplets et/ou des montants de mauvais sens. Ces éléments ont été rejetés.')

    else if lBoTransacOk then
      PGIInfo('Les pièces et le ficher CFONB ont été correctement générés.', Ecran.Caption) ;
    end
  else if lBoTransacOk
    then PGIInfo('Les pièces comptables ont été correctement générées.', Ecran.Caption)
    else PGIInfo('Le traitement n''a pas pu aboutir.', Ecran.Caption) ;

end;


function TOF_CPGENERENCADECA.ConstitueOrigineOpti: Tob;
Var lStReq : String ;
begin
  Result := TOB.Create( 'TOB_ORIGINE', nil, -1 ) ;

  lStReq := 'SELECT ##TOP ' + IntToStr( FMaxLignes ) + '## ' + GetChampsOpti + ', ' + TPieceCompta.GetChampsAuxi + ', ' + TPieceCompta.GetChampsGene
               + ' FROM ECRITURE'
                   + ' LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE'
                   + ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL '
               + GetConditions ;

  result.LoadDetailFromSQL( lStReq );

end;

procedure TOF_CPGENERENCADECA.SetCriteresOpti;
var lStWhere : string ;
begin
  SetChampsInvisibles ;
  SetControlText('XX_WHEREMONTANT', '' ) ;
  SetControlText('E_ETATLETTRAGE', '' ) ;

  lStWhere := GetControlText('XX_WHERE') ;
  if Trim(lStWhere) = ''
    then lStWhere := 'E_ETATLETTRAGE="AL"'
    else lStWhere := lStWhere + ' AND E_ETATLETTRAGE="AL" ' ;

  if GetControlText('SOLDEPOSITIF')='X' then
    if GetControlText('FTIC')<>'X' then
      lStWhere := lStWhere + ' AND T_TOTALDEBIT>T_TOTALCREDIT ' ;

  SetControlText( 'XX_WHERE', lStWhere ) ;

end;

procedure TOF_CPGENERENCADECA.InitScenarioOpti;
begin

  // Création de la TOB
  if Assigned(TobScenario) then
    FreeAndNil( TOBScenario ) ;

  TOBScenario := TOB.Create('CPARAMGENER', nil, -1) ;
  AjouteChampsSuppScenario( TOBScenario ) ;

  TobScenario.PutValue('CPG_TYPEENCADECA', 'PRE') ;
  TobScenario.PutValue('CPG_FLUXENCADECA', Flux ) ;

  TobScenario.PutValue('CPG_JOURNAL', GetControlText('CPGJOURNAL') ) ;
  TobScenario.PutValue('CPG_GENERAL', GetControlText('CPGGENERAL') ) ;
  TobScenario.PutValue('CPG_MODEPAIEMENT', GetControlText('CPGMODEPAIE') ) ;

  TobScenario.PutValue('NOMLOT', 'TEMP/' + V_PGI.User + '/' + FormatDateTime('yyyymmdd',Date) )  ;

  TobScenario.AddChampSupValeur('TRESOSYNCHRO', ets_Nouveau) ;

end;

procedure TOF_CPGENERENCADECA.InitMulOpti;
begin
  SetControlVisible('PAVANCES2', False) ;
  SetControlVisible('OPERATION', False) ;
  SetControlVisible('TOPERATION', False) ;
  SetControlVisible('SCENARIO', False) ;
  SetControlVisible('TSCENARIO', False) ;
  SetControlVisible('ETABLISSEMENTPAYEUR', False) ;
  SetControlVisible('TETABLISSEMENTPAYEUR', False) ;
  SetControlVisible('MULTIDOSSIER', False) ;
  SetControlVisible('TMULTIDOSSIER', False) ;

  SetControlVisible('PANELSCENARIO', False) ;
//  SetControlProperty('PANELSCENARIO', 'WIDTH', 0 ) ;

{ // FQ 18885 : Ajout de la gestion des TIC / TID
  SetControlVisible('FTIC', False) ;
  FTic.Checked := False ;
  FTICChanged( nil ) ;
}
end;

function TOF_CPGENERENCADECA.VerifieParametresOpti: Boolean;
begin
  Result := False ;

  // Test sur Journal
  if not LookUpValueExist( GetControl('CPGJournal') ) then
    begin
    PGIError('Vous devez renseigner un journal de génération.', Ecran.Caption ) ;
    SetActiveTabSheet('PCRITERE') ;
    SetFocusControl('CPGJOURNAL') ;
    Exit ;
    end ;

  // Test sur compte de génération
  if not LookUpValueExist( GetControl('CPGGeneral',True) ) then
    begin
    PGIError('Vous devez renseigner un compte de génération.', Ecran.Caption ) ;
    SetActiveTabSheet('PCRITERE') ;
    SetFocusControl('CPGGENERAL') ;
    Exit ;
    end ;

  // Test sur compte de génération
  if not LookUpValueExist( GetControl('CPGMODEPAIE',True) ) then
    begin
    PGIError('Vous devez renseigner un mode de paiement.', Ecran.Caption ) ;
    SetActiveTabSheet('PCRITERE') ;
    SetFocusControl('CPGMODEPAIE') ;
    Exit ;
    end ;

  // Test journal de substitution pour les pièces sans contrepartie
  if ( GetParamSocSecur( 'SO_CPJALENCADECA', '' ) = '' ) then
    begin
    PGIInfo('Vous devez renseigner le journal de substitution (paramètre sociétés - suivi-tiers) pour pouvoir utiliser la fonction.', Ecran.Caption) ;
    exit ;
    end ;

  // Test journal de substitution pour les pièces sans contrepartie
  if ( GetParamSocSecur( 'SO_CPCPTENCADECA', '' ) = '' ) then
    begin
    PGIInfo('Vous devez renseigner le compte de substitution (paramètre sociétés - suivi-tiers) pour pouvoir utiliser la fonction.', Ecran.Caption) ;
    exit ;
    end ;

  // Tests Date
  if not DatesOk then Exit ;

  if FMaxLignes = 0 then
    begin
    PGIInfo('Vous devez saisir le nombre maximum d''écheances à traiter par pièce générée.', Ecran.Caption) ;
    SetActiveTabSheet('PCRITERE') ;
    SetFocusControl('NBLIGNES') ;
    exit ;
    end ;

  if FMaxLignes > 1000 then
    begin
    if PGIAsk('Vous avez demandé le traitement de plus de 1000 échéances par pièce. Confirmez-vous le traitement ?', Ecran.Caption) <> mrYes then
      exit ;
    end ;

  // Comptage et validation
  if not CompteEche then
    exit ;

  // Tout est ok
  Result := True ;

end;

procedure TOF_CPGENERENCADECA.CPGJournalChange(Sender: TObject);
var Q : TQuery;
begin

  SetControlEnabled ('CPGGENERAL', True ) ;
  SetControlText    ('CPGGENERAL', '' ) ;

  if (GetControlText('CPGJOURNAL')<>'') then
  begin
    Q := OpenSQL('SELECT J_NATUREJAL, J_EFFET, J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="' + GetControlText('CPGJOURNAL') + '" ', TRUE);
    if not Q.Eof then
      begin
      SetControlText('CPGGENERAL', Q.FindField('J_CONTREPARTIE').AsString ) ;
      SetControlEnabled('CPGGENERAL', not ((Q.FindField('J_NATUREJAL').AsString = 'BQE')
                                      or (Q.FindField('J_NATUREJAL').AsString = 'CAI')
                                      or (Q.FindField('J_EFFET').AsString = 'X') )       ) ;
      end;
    Ferme(Q);
  end;

end;

function TOF_CPGENERENCADECA.TraitementLotOpti( vTobOrigine: TOB ; vBoRecap : Boolean = False ): boolean;
var i        : integer ;
    lTobGene : Tob ;
    lTobEcr  : Tob ;
    lTobEch  : Tob ;
    lDebit   : Double ;
    lCredit  : Double ;
    lTotalD  : Double ;
    lTotalC  : Double ;
    lTLCFONB : TList ;
begin

  result := False ;
  if vTobOrigine.Detail.Count = 0 then Exit ;

  lTotalD  := 0 ;
  lTotalC  := 0 ;
  lTobGene := Tob.create('$PIECE', nil, -1 ) ;

  // Recup info Gene
  RecupInfoCptGene ;
  lTlCfonb := TList.Create;


  // ******************************
  // ***  CREATION LIGNES REGLT ***
  // ******************************
  lTobEch := nil ;
  for i := 1 to vTobOrigine.detail.count do
    begin

    lTobEch  := vTobOrigine.detail[ i - 1 ] ;
    lDebit   := 0 ;
    lCredit  := 0 ;

    // Création ligne
    lTobEcr := Tob.Create('ECRITURE', lTobGene, -1 ) ;
    CPutDefautEcr( lTobEcr ) ;
    InitEnteteEcr( lTobEcr ) ;
    RecopieEcheances( lTobEcr, lTobEch, False ) ;

    // Montants
    MajCumulsEncaDeca( lTobEch, lDebit, lCredit ) ;
    CSetMontants( lTobEcr, lCredit, lDebit , Fdev, True ) ;

    // Analytique
    if lTobEch.GetString('E_ANA') = 'X' then
      VentilerTob(lTobEcr, '', 0, FDev.Decimale, False ) ;

    // Mise en place du lettrage
    if (not vBoRecap) and ( lTobEch.GetString('E_ECHE')='X' ) then
      RenseigneLettrage( lTobEcr, lTobEch ) ;

    // CFONB
    if (not vBoRecap) and FBoCfonb then
      begin
      RenseigneCFONB( lTobEcr, lTlCfonb ) ;
      lTobEcr.AddChampSupValeur('T_LIBELLE',  lTobEch.GetValue('T_LIBELLE') ) ;
      lTobEcr.AddChampSupValeur('G_LIBELLE',  lTobEch.GetValue('G_LIBELLE') ) ;
      end ;

    // cumuls des montants
    MajCumulsEncaDeca( lTobEch, lTotalD, lTotalC ) ;

    end ;

  // ************************************
  // ***  CREATION LIGNE CONTREPARTIE ***
  // ************************************
    // Création ligne
    lTobEcr := Tob.Create('ECRITURE', lTobGene, -1 ) ;
    CPutDefautEcr( lTobEcr ) ;
    InitEnteteEcr( lTobEcr ) ;
    RecopieEcheances( lTobEcr, lTobEch, True ) ;
    lTobEcr.PutValue('E_TRESOSYNCHRO', TobScenario.GetValue('TRESOSYNCHRO') ) ;
    if not vBoRecap then
      lTobEcr.PutValue('E_NOMLOT',       TobScenario.GetValue('NOMLOT') ) ;

    // Montants
    CSetMontants( lTobEcr, lTotalD, lTotalC, Fdev, True ) ;

    // Analytique
    if FTobCptGene.GetString('G_VENTILABLE') = 'X' then
      VentilerTob(lTobEcr, '', 0, FDev.Decimale, False ) ;

  // *****************
  // ***  MAJ BASE ***
  // *****************

  // Numéro de pièce
  result := EnregistrePiece( lTobGene, vBoRecap ) ;
  
  // MAJ des échéances (pour lettrage)
  if not vBoRecap then
    UpdateEcheance( vTobOrigine ) ;

  // ***************
  // ***  EXPORT ***
  // ***************
  if (not vBoRecap) and FBoCfonb then
    ExportCFONBBatch( TobScenario, lTlCfonb ) ;

  // *********************
  // ***  MAJ MESSAGES ***
  // *********************

  // Récup dernière ligne pour affichage
  lTobEcr.ChangeParent( FTobRefPiece, -1 ) ;
  AffichePieces ;
  Application.ProcessMessages ;

  // Libération mémoire
  FreeAndNil( lTobGene ) ;
  if Assigned( FTobLETTRAGE ) then
    FreeAndNil( FTobLETTRAGE ) ;
  if Assigned( lTlCfonb ) then
    FreeAndNil( lTlCfonb ) ;

end;

procedure TOF_CPGENERENCADECA.RecopieEcheances(vTobEcr, vTobOrig: Tob ; vBoContrepartie : Boolean );
begin

  if vBoContrepartie then
    begin
    // Comptes
    vTobEcr.PutValue('E_GENERAL' ,         TobScenario.GetValue('CPG_GENERAL') ) ;
    vTobEcr.PutValue('E_AUXILIAIRE' ,      '' );
    vTobEcr.PutValue('E_CONTREPARTIEGEN',  vTobOrig.GetValue('E_GENERAL') ) ;
    vTobEcr.PutValue('E_CONTREPARTIEAUX',  '' ) ;

    // Paiement
    if ( TobScenario.GetValue('FORCERDATEECHE')='X' ) then
      vTobEcr.PutValue('E_DATEECHEANCE',   TobScenario.GetValue('DATEECHE') ) ;

    if FTobCptGene.GetString('G_POINTABLE')='X' then
      begin
      CRemplirInfoPointage( vTobEcr ) ;
      vTobEcr.PutValue('E_MODEPAIE' ,      TobScenario.GetValue('CPG_MODEPAIEMENT') );
      vTobEcr.PutValue('E_DATEECHEANCE',   vTobOrig.GetValue('E_DATEECHEANCE') );
      vTobEcr.PutValue('E_ORIGINEPAIEMENT',vTobOrig.GetValue('E_ORIGINEPAIEMENT') );
      vTobEcr.PutValue('E_DATEVALEUR' ,    vTobOrig.GetValue('E_DATEECHEANCE') ) ;
      end
    else if FTobCptGene.GetString('G_LETTRABLE')='X' then
      begin
      CRemplirInfoLettrage( vTobEcr ) ;
      if FTobCptGene.GetString('G_NATUREGENE')='DIV' then
        vTobEcr.PutValue( 'E_TRESOSYNCHRO', 'RIE' ) ;
      vTobEcr.PutValue('E_MODEPAIE' ,      TobScenario.GetValue('CPG_MODEPAIEMENT') );
      vTobEcr.PutValue('E_DATEECHEANCE',   vTobOrig.GetValue('E_DATEECHEANCE') );
      vTobEcr.PutValue('E_ORIGINEPAIEMENT',vTobOrig.GetValue('E_ORIGINEPAIEMENT') );
      vTobEcr.PutValue('E_DATEVALEUR' ,    vTobOrig.GetValue('E_DATEECHEANCE') ) ;
      end
    else CSupprimerInfoLettrage( vTobEcr ) ;

    end
  else
    begin
    // Comptes
    vTobEcr.PutValue('E_GENERAL' ,        vTobOrig.GetValue('E_GENERAL') );
    vTobEcr.PutValue('E_AUXILIAIRE' ,     vTobOrig.GetValue('E_AUXILIAIRE') );
    vTobEcr.PutValue('E_CONTREPARTIEGEN', TobScenario.GetValue('CPG_GENERAL') ) ;
    vTobEcr.PutValue('E_CONTREPARTIEAUX', '' ) ;

    // Paiement
    if vTobOrig.GetString('E_ECHE') = 'X' then
      CRemplirInfoLettrage( vTobEcr ) ;

    vTobEcr.PutValue('E_MODEPAIE' ,      TobScenario.GetValue('CPG_MODEPAIEMENT') );
    vTobEcr.PutValue('E_DATEECHEANCE',   vTobOrig.GetValue('E_DATEECHEANCE') );
    vTobEcr.PutValue('E_ORIGINEPAIEMENT',vTobOrig.GetValue('E_ORIGINEPAIEMENT') );
    vTobEcr.PutValue('E_DATEVALEUR' ,    vTobOrig.GetValue('E_DATEECHEANCE') ) ;
    if ( TobScenario.GetValue('FORCERDATEECHE')='X' ) then
      vTobEcr.PutValue('E_DATEECHEANCE',   TobScenario.GetValue('DATEECHE') ) ;
    end ;

  // Divers ;
  vTobEcr.PutValue('E_REFINTERNE',     vTobOrig.GetValue('E_REFINTERNE') ) ;
  vTobEcr.PutValue('E_LIBELLE',        vTobOrig.GetValue('E_LIBELLE') ) ;
  vTobEcr.PutValue('E_RIB',            vTobOrig.GetValue('E_RIB') ) ;
  vTobEcr.PutValue('E_TVAENCAISSEMENT',vTobOrig.GetValue('E_TVAENCAISSEMENT') ) ;
  vTobEcr.PutValue('E_REGIMETVA',      vTobOrig.GetValue('E_REGIMETVA') ) ;
  vTobEcr.PutValue('E_TVA',            vTobOrig.GetValue('E_TVA') ) ;
  vTobEcr.PutValue('E_TPF',            vTobOrig.GetValue('E_TPF') ) ;
  vTobEcr.PutValue('E_REFEXTERNE',     vTobOrig.GetValue('E_REFEXTERNE') ) ;
  vTobEcr.PutValue('E_DATEREFEXTERNE', vTobOrig.GetValue('E_DATEREFEXTERNE') ) ;
  vTobEcr.PutValue('E_REFLIBRE',       vTobOrig.GetValue('E_REFLIBRE') ) ;
  vTobEcr.PutValue('E_AFFAIRE',        vTobOrig.GetValue('E_AFFAIRE') ) ;
  vTobEcr.PutValue('E_QTE1',           vTobOrig.GetValue('E_QTE1') ) ;
  vTobEcr.PutValue('E_QTE2',           vTobOrig.GetValue('E_QTE2') ) ;
  vTobEcr.PutValue('E_QUALIFQTE1',     vTobOrig.GetValue('E_QUALIFQTE1') ) ;
  vTobEcr.PutValue('E_QUALIFQTE2',     vTobOrig.GetValue('E_QUALIFQTE2') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE0',    vTobOrig.GetValue('E_LIBRETEXTE0') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE1',    vTobOrig.GetValue('E_LIBRETEXTE1') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE2',    vTobOrig.GetValue('E_LIBRETEXTE2') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE3',    vTobOrig.GetValue('E_LIBRETEXTE3') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE4',    vTobOrig.GetValue('E_LIBRETEXTE4') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE5',    vTobOrig.GetValue('E_LIBRETEXTE5') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE6',    vTobOrig.GetValue('E_LIBRETEXTE6') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE7',    vTobOrig.GetValue('E_LIBRETEXTE7') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE8',    vTobOrig.GetValue('E_LIBRETEXTE8') ) ;
  vTobEcr.PutValue('E_LIBRETEXTE9',    vTobOrig.GetValue('E_LIBRETEXTE9') ) ;
  vTobEcr.PutValue('E_TABLE0',         vTobOrig.GetValue('E_TABLE0') ) ;
  vTobEcr.PutValue('E_TABLE1',         vTobOrig.GetValue('E_TABLE1') ) ;
  vTobEcr.PutValue('E_TABLE2',         vTobOrig.GetValue('E_TABLE2') ) ;
  vTobEcr.PutValue('E_TABLE3',         vTobOrig.GetValue('E_TABLE3') ) ;
  vTobEcr.PutValue('E_LIBREDATE',      vTobOrig.GetValue('E_LIBREDATE') ) ;
  vTobEcr.PutValue('E_LIBREBOOL0',     vTobOrig.GetValue('E_LIBREBOOL0') ) ;
  vTobEcr.PutValue('E_LIBREBOOL1',     vTobOrig.GetValue('E_LIBREBOOL1') ) ;
  vTobEcr.PutValue('E_LIBREMONTANT0',  vTobOrig.GetValue('E_LIBREMONTANT0') ) ;
  vTobEcr.PutValue('E_LIBREMONTANT1',  vTobOrig.GetValue('E_LIBREMONTANT1') ) ;
  vTobEcr.PutValue('E_LIBREMONTANT2',  vTobOrig.GetValue('E_LIBREMONTANT2') ) ;
  vTobEcr.PutValue('E_LIBREMONTANT3',  vTobOrig.GetValue('E_LIBREMONTANT3') ) ;
  vTobEcr.PutValue('E_CONSO',          vTobOrig.GetValue('E_CONSO') ) ;

end;

procedure TOF_CPGENERENCADECA.RecupInfoCptGene;
begin

  if not Assigned( FTobCptGene ) then
    FTobCptGene := Tob.Create('GENERAUX', nil, -1) ;

  if FTobCptGene.GetString('G_GENERAL') <> TobScenario.GetString('CPG_GENERAL') then
    begin
    FTobCptGene.PutValue('G_GENERAL', TobScenario.GetString('CPG_GENERAL') ) ;
    FTobCptGene.LoadDB ;
    end ;

end;

procedure TOF_CPGENERENCADECA.RenseigneCFONB( vTobEcr: Tob ; vTLCfonb : Tlist );
begin
  if not FBoCFONB then Exit ;
  vTobEcr.PutValue('E_CFONBOK',  'X' ) ;
  vTobEcr.PutValue('E_NUMCFONB', TobScenario.GetString('CFONB_NUMCFONB') ) ;
  vTlCFONB.Add( vTobEcr ) ;
end;

procedure TOF_CPGENERENCADECA.RenseigneLettrage(vTobEcr, vTobEch: Tob);
var lTobAux : Tob ;
    lCodeL  : string ;
    lAux    : string ;
    Q       : TQuery ;
begin

  if not Assigned(FTobLETTRAGE) then
    FTobLETTRAGE := Tob.Create('$CODELETT', nil, -1) ;

  if FTIC.Checked // FQ 18885 SBO 31/10/2006 : Ajout gestin des tic tid
    then lAux    := vTobEch.GetString('E_GENERAL')
    else lAux    := vTobEch.GetString('E_AUXILIAIRE') ;
  lCodeL  := '' ;

  lTobAux := FTobLETTRAGE.FindFirst(['E_AUXILIAIRE'], [lAux], False) ;
  if lTobAux = nil then
    begin
    // E_LETTRAGE
    if FTIC.Checked // FQ 18885 SBO 31/10/2006 : Ajout gestin des tic tid
     then Q:=OpenSQL('Select G_DERNLETTRAGE AS DERNLETTRAGE from GENERAUX Where G_GENERAL="'+lAux+'"',True)
     else Q:=OpenSQL('Select T_DERNLETTRAGE AS DERNLETTRAGE from TIERS Where T_AUXILIAIRE="'+lAux+'"',True) ;
    lCodeL := '' ;
    if Not Q.EOF then
      begin
      lCodeL:=Q.FindField('DERNLETTRAGE').AsString ;
      Ferme (Q);
      lCodeL:=CodeSuivant(lCodeL) ;
      end
    else Ferme (Q);


    if lCodeL = '' then Exit ;

    if FTIC.Checked // FQ 18885 SBO 31/10/2006 : Ajout gestin des tic tid
      then ExecuteSQL('UPDATE GENERAUX SET G_DERNLETTRAGE="'+lCodeL+'" WHERE G_GENERAL = "'+lAux+'"')
      else ExecuteSQL('UPDATE TIERS SET T_DERNLETTRAGE="'+lCodeL+'" WHERE T_AUXILIAIRE = "'+lAux+'"') ;


    lTobAux := Tob.Create('AUX', FTobLETTRAGE, -1 ) ;
    lTobAux.AddChampSupValeur('E_AUXILIAIRE', lAux ) ;
    // INIT DATES
    lTobAux.AddChampSupValeur('E_DATEPAQUETMIN', IDate2099 ) ;
    lTobAux.AddChampSupValeur('E_DATEPAQUETMAX', IDate1900 ) ;
    lTobAux.AddChampSupValeur('E_LETTRAGE', lCodeL ) ;
    // E_ETAT
    if FDev.Code <> V_PGI.DevisePivot then
      begin
      lTobAux.AddChampSupValeur('E_ETAT', '--X0AM0000' ) ;
      lTobAux.AddChampSupValeur('E_LETTRAGEDEV', 'X' )
      end
    else
      begin
      lTobAux.AddChampSupValeur('E_ETAT', '---0AM0000' ) ;
      lTobAux.AddChampSupValeur('E_LETTRAGEDEV', '-' )
      end ;
    end
  else
    lCodeL := lTobAux.GetValue('E_LETTRAGE') ;

  if (lTobAux = nil) or (lCodeL = '') then Exit ;

  // Calcul dates paquets
  if vTobEch.GetValue('E_DATECOMPTABLE') < lTobAux.GetValue('E_DATEPAQUETMIN') then
    lTobAux.PutValue('E_DATEPAQUETMIN', vTobEch.GetValue('E_DATECOMPTABLE') ) ;
  if vTobEcr.GetValue('E_DATECOMPTABLE') < lTobAux.GetValue('E_DATEPAQUETMIN') then
    lTobAux.PutValue('E_DATEPAQUETMIN', vTobEcr.GetValue('E_DATECOMPTABLE') ) ;
  if vTobEch.GetValue('E_DATECOMPTABLE') > lTobAux.GetValue('E_DATEPAQUETMAX') then
    lTobAux.PutValue('E_DATEPAQUETMAX', vTobEch.GetValue('E_DATECOMPTABLE') ) ;
  if vTobEcr.GetValue('E_DATECOMPTABLE') > lTobAux.GetValue('E_DATEPAQUETMAX') then
    lTobAux.PutValue('E_DATEPAQUETMAX', vTobEcr.GetValue('E_DATECOMPTABLE') ) ;

  vTobEch.AddChampSupValeur('LETTRAGE', 'X' );

end;

function  TOF_CPGENERENCADECA.EnregistrePiece( vTobGene: Tob ; vBoRecap : Boolean = False ) : boolean ;
var i       : integer ;
    lTobEcr : Tob ;
    lDateM  : TDateTime ;
    lNumP   : integer ;
    lTobLet : Tob ;
    lAux    : string ;
{
    lPC     : TPieceCompta ;
    lInfoEcr : TInfoEcriture ;}
begin

//  result := False ;

  // Uniformisations des infos pièces tel que numéro, date modif, nature pièce...
  lNumP  := GetNewNumJal( TobScenario.GetValue('CPG_JOURNAL'), True, TobScenario.GetValue('DATECOMPTABLE') ) ;
  lDateM := NowH ;

  For i := 1 to vTObGene.Detail.Count do
    begin

    lTobEcr := vTObGene.Detail[ i-1 ] ;

    // infos pièces
    lTobEcr.PutValue('E_NUMEROPIECE', lNumP   ) ;
    lTobEcr.PutValue('E_DATEMODIF',   lDateM  ) ;
    lTobEcr.PutValue('E_NUMLIGNE',    i       ) ;

    // Lettrage des échéances
    if not vBoRecap and (i < vTObGene.Detail.Count) then
      begin
      if lTobEcr.GetValue('E_ECHE')='X' then
        begin
        if FTIC.Checked // FQ 18885 SBO 31/10/2006 : Ajout gestin des tic tid
          then lAux    := lTobEcr.GetString('E_GENERAL')
          else lAux    := lTobEcr.GetString('E_AUXILIAIRE') ;
        lTobLet := FTobLETTRAGE.FindFirst(['E_AUXILIAIRE'], [lAux], False) ;
        if lTobLet <> nil then
          begin
          lTobEcr.PutValue('E_ETAT',           lTobLet.GetValue('E_ETAT') ) ;
          lTobEcr.PutValue('E_LETTRAGE',       lTobLet.GetValue('E_LETTRAGE') ) ;
          lTobEcr.PutValue('E_LETTRAGEDEV',    lTobLet.GetValue('E_LETTRAGEDEV') ) ;
          lTobEcr.PutValue('E_ETATLETTRAGE',   'TL' ) ;
          lTobEcr.PutValue('E_DATEPAQUETMIN',  lTobLet.GetValue('E_DATEPAQUETMIN') ) ;
          lTobEcr.PutValue('E_DATEPAQUETMAX',  lTobLet.GetValue('E_DATEPAQUETMAX') ) ;
          lTobEcr.PutValue('E_COUVERTURE',     lTobEcr.GetDouble('E_DEBIT')+lTobEcr.GetDouble('E_CREDIT') ) ;
          lTobEcr.PutValue('E_COUVERTUREDEV',  lTobEcr.GetDouble('E_DEBITDEV')+lTobEcr.GetDouble('E_CREDITDEV') ) ;
          end ;
        end ;
      end ;

    end ;

  // TEST sur TPIECECOMPTA
{  lInfoEcr := TInfoEcriture.Create( '' ) ;
  lPC      := TPieceCompta.CreerPiece( lInfoEcr ) ;
  lPC.LoadFromTob( vTobGene ) ;
  if not lPC.IsValidPiece then Exit ;

  // Récupération des filles de la Pièce
  for i:=0 to lPC.detail.count - 1 do
    lPC.detail[ i ].ChangeParent( vTobGene, -1 ) ;
}
  // enregistrement du règlement
  result := vTobGene.InsertDB( nil ) ;

  // Maj des soldes
  if result then
    MajSoldesEcritureTOB ( vTobGene , True ) ;

end;

procedure TOF_CPGENERENCADECA.InitEnteteEcr( vTobEcr: Tob );
begin
  vTobEcr.PutValue('E_JOURNAL',       TobScenario.GetValue('CPG_JOURNAL') ) ;
  vTobEcr.PutValue('E_NATUREPIECE',   TobScenario.GetValue('NATUREPIECE') ) ;
  vTobEcr.PutValue('E_DEVISE',        FDev.Code ) ;
  vTobEcr.PutValue('E_QUALIFPIECE',   'N' ) ;
  vTobEcr.PutValue('E_ECRANOUVEAU',   'N' ) ;
  vTobEcr.PutValue('E_MODESAISIE',    '-' ) ;
  CRemplirDateComptable( vTobEcr,     TobScenario.GetValue('DATECOMPTABLE') ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGENERENCADECA.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPGENERENCADECA.UpdateEcheance(vTobOrigine: Tob);
var i       : integer ;
    lTobEcr : Tob ;
    lDateM  : TDateTime ;
    lTobLet : Tob ;
    lSQL    : string ;
    lCouv   : Double ;
    lCouvD  : Double ;
    lAux    : String ;
begin

  lDateM := NowH ;

  For i := 1 to vTobOrigine.Detail.Count do
    begin

    lTobEcr := vTobOrigine.Detail[ i-1 ] ;

    if lTobEcr.GetNumChamp('LETTRAGE') < 0 then Continue ;


    if lTobEcr.GetValue('E_ECHE')='X' then
      begin

      if FTIC.Checked // FQ 18885 SBO 31/10/2006 : Ajout gestin des tic tid
        then lAux    := lTobEcr.GetString('E_GENERAL')
        else lAux    := lTobEcr.GetString('E_AUXILIAIRE') ;
      lTobLet := FTobLETTRAGE.FindFirst(['E_AUXILIAIRE'], [lAux], False) ;
      if lTobLet <> nil then
        begin
{
        lTobEcr.PutValue('E_ETAT',           lTobLet.GetValue('E_ETAT') ) ;
        lTobEcr.PutValue('E_LETTRAGE',       lTobLet.GetValue('E_LETTRAGE') ) ;
        lTobEcr.PutValue('E_LETTRAGEDEV',    lTobLet.GetValue('E_LETTRAGEDEV') ) ;
        lTobEcr.PutValue('E_ETATLETTRAGE',   'TL' ) ;
        lTobEcr.PutValue('E_DATEPAQUETMIN',  lTobLet.GetValue('E_DATEPAQUETMIN') ) ;
        lTobEcr.PutValue('E_DATEPAQUETMAX',  lTobLet.GetValue('E_DATEPAQUETMAX') ) ;
        lTobEcr.PutValue('E_COUVERTURE',     lTobEcr.GetDouble('E_DEBIT')+lTobEcr.GetDouble('E_CREDIT') ) ;
        lTobEcr.PutValue('E_COUVERTUREDEV',  lTobEcr.GetDouble('E_DEBITDEV')+lTobEcr.GetDouble('E_CREDITDEV') ) ;
}
        lCouv  := lTobEcr.GetDouble('E_CREDIT') + lTobEcr.GetDouble('E_DEBIT') ;
        lCouvD := lTobEcr.GetDouble('E_CREDITDEV') + lTobEcr.GetDouble('E_DEBITDEV') ;

        lSQL := 'UPDATE ECRITURE SET E_LETTRAGE="'      + lTobLet.GetValue('E_LETTRAGE')                    + '", '
                                  + 'E_ETATLETTRAGE="TL", '
                                  + 'E_DATEPAQUETMIN="' + USDATETIME( lTobLet.GetValue('E_DATEPAQUETMIN') ) + '", '
                                  + 'E_DATEPAQUETMAX="' + USDATETIME( lTobLet.GetValue('E_DATEPAQUETMAX') ) + '", '
                                  + 'E_COUVERTURE='     + StrfPoint(lCouv)+', '
                                  + 'E_COUVERTUREDEV='  + StrfPoint(lCouvD)+', '
                                  + 'E_LETTRAGEDEV="'   + lTobLet.GetValue('E_LETTRAGEDEV') + '", '
                                  + 'E_ETAT="'          + lTobLet.GetValue('E_ETAT') + '", '
                                  + 'E_DATEMODIF="'     + USTime( lDateM ) + '" ' ;
        if lTobEcr.GetValue('E_TRESOSYNCHRO') <> ets_Rien then
         lSQL := lSQL + ', E_TRESOSYNCHRO = "' + ets_Lettre + '" ' ;

         lSQL := lSQL + ' WHERE E_JOURNAL="'       + lTobEcr.GetString( 'E_JOURNAL')     + '"' +
                    ' AND E_EXERCICE="'      + lTobEcr.GetString( 'E_EXERCICE')    + '"' +
                    ' AND E_DATECOMPTABLE="' + UsDateTime( lTobEcr.GetDateTime( 'E_DATECOMPTABLE') ) + '"' +
                    ' AND E_NUMEROPIECE='    + lTobEcr.GetString( 'E_NUMEROPIECE') +
                    ' AND E_QUALIFPIECE="'   + lTobEcr.GetString( 'E_QUALIFPIECE') + '"' +
                    ' AND E_NUMLIGNE='       + lTobEcr.GetString( 'E_NUMLIGNE') +
                    ' AND E_NUMECHE='        + lTobEcr.GetString( 'E_NUMECHE')  ;

        ExecuteSQL( lSQL ) ;
        end ;
    end ;
  end ;


end;

procedure TOF_CPGENERENCADECA.ClickValide;
begin
  if FBoVersionOpti
    then ClickValideOpti
    else ClickValideOptiBis ;
end;

procedure TOF_CPGENERENCADECA.ClickValideOptiBis;
var lInLot         : integer ;
    lStJal         : string ;
    lStCpt         : string ;
    lInCodeErr     : integer ;
    lTobOrigine    : TOB ;
    lStLot         : string ;
begin

  FMaxLignes    := ValeurI(GetControlText('NBLIGNES'));
  lStJal        := TobScenario.Getstring('CPG_JOURNAL' ) ;
  lStCpt        := TobScenario.Getstring('CPG_GENERAL' ) ;
  lInCodeErr    := CGE_PASERREUR ;

  // Maj des champs supplémentaires de paramètrage
  CompleteTobScenario ;
  TobScenario.PutValue('NOMLOT', 'TEMP/' + V_PGI.User + '/' + FormatDateTime('yyyymmdd',Date) )  ;

  // Test d'usage
  if not VerifieParametres then Exit ;

  // Vide Grille
  if Assigned(FTobRefPiece)
    then FTobRefPiece.ClearDetail
    else FTobRefPiece  := Tob.Create('$REFPIECE', nil, -1 ) ;
  VideGrille ;

  // Recup info CFONB
  if FBoCfonb then
    if not ParamCFONB( TobScenario ) then Exit ;

  try

    lInLot := 0 ;

    repeat

     Inc( lInLot ) ;
     // indicateurs CFONB
     if FBoCfonb and (lInLot > 1) then
       TobScenario.PutValue('CFONB_PREMIERLOT', '-' ) ;
     // Message
     SetControlText('MESSAGE', 'Génération du lot n° ' + IntToStr( lInLot ) + ' / ' + IntToStr ( GetMaxLot ) + '...' ) ;
     Application.ProcessMessages ;

     // Récup des échéances d'origine
     if Assigned( lTobOrigine ) then
       begin
       lTobOrigine.ClearDetail ;
       FreeAndNil( lTobOrigine ) ;
       end ;
     lTobOrigine := ConstitueOrigineOpti ;

     // traitement des lots :
     if lTobOrigine.detail.count > 0 then
       begin

       // Si génération globalise multi-lot, positionnement sur jal et cpt de substitution
       if lTobOrigine.detail.count >= FMaxLignes then
         begin
         if lInLot = 1 then
           begin
           TobScenario.PutValue('CPG_JOURNAL', GetParamSocSecur( 'SO_CPJALENCADECA', '') ) ;
           TobScenario.PutValue('CPG_GENERAL', GetParamSocSecur( 'SO_CPCPTENCADECA', '') ) ;
           TobScenario.PutValue('TRESOSYNCHRO', ets_Rien ) ;
           end ;
         end
       else
         begin
         if lInLot = 1 then
           TobScenario.PutValue( 'NOMLOT',  ''  ) ;
         if FBoCfonb then
           TobScenario.PutValue('CFONB_DERNIERLOT', 'X' ) ; // indicateurs CFONB
         end ;

       // Traitement lot
       lInCodeErr := ExecuteEncaDeca( lTobOrigine , TobScenario, FTobRefPiece ) ;
       if (lInCodeErr = CGE_PASERREUR) then
         begin
         AffichePieces ;
         Application.ProcessMessages ;
         end ;

       end ;

     until ( lTobOrigine.Detail.count = 0 ) ;

     // Traitement pièce finale
     if (lInLot > 1) and (TobScenario.GetValue( 'NOMLOT' ) <> '') then
       begin
       TobScenario.PutValue('CPG_JOURNAL', lStJal ) ;
       TobScenario.PutValue('CPG_GENERAL', lStCpt ) ;
       TobScenario.PutValue('CPG_CFONBEXPORT', '-' ) ;
       TobScenario.PutValue('TRESOSYNCHRO', ets_Nouveau ) ;
       TobScenario.PutValue('SANSLETTRAGE', 'X') ;

       SetControlText('MESSAGE', 'Génération de la pièce finale...' ) ;
       Application.ProcessMessages ;

       if Assigned( lTobOrigine ) then
         begin
         lTobOrigine.ClearDetail ;
         FreeAndNil( lTobOrigine ) ;
         end ;

       lStLot       := TobScenario.GetValue( 'NOMLOT' ) ;
       lTOBOrigine  := ConstitueOrigineFinal( lStLot ) ;
       TobScenario.PutValue( 'NOMLOT',  ''  ) ;

       lInCodeErr := ExecuteEncaDeca( lTobOrigine , TobScenario, FTobRefPiece ) ;
       if (lInCodeErr=CGE_PASERREUR) then
         begin
         AffichePieces ;
         Application.ProcessMessages ;
         end ;

      // Effacement du nom de lot temporaire
      ExecuteSQL('UPDATE ECRITURE SET E_NOMLOT="" WHERE E_NOMLOT="' + lStLot + '"' ) ;

      end ;


    except
      on E: Exception do
        PgiError('Erreur lors du traitement : ' + E.Message, Ecran.caption ) ;
    end ;

  if Assigned( lTobOrigine ) then
     begin
     lTobOrigine.ClearDetail ;
     FreeAndNil( lTobOrigine ) ;
     end ;

  // Messages
  SetControlText('MESSAGE', 'Liste des pièces générées' ) ;

  if FBoCfonb then
    begin

    if (TobScenario.GetNumChamp('CFONB_ERREUR') > 0) and (TobScenario.GetValue('CFONB_ERREUR') > 0 )  then
      begin
      if TobScenario.GetValue('CFONB_ERREUR') = 16 then
        PgiInfo('Certains lots n''ont pas pu être exportés !', Ecran.Caption) ;
      end

    else if (TobScenario.GetNumChamp('CFONB_EXISTEMAUVAIS') > 0) and (TobScenario.GetValue('CFONB_EXISTEMAUVAIS') = 'X' )  then
        PgiInfo('Vous avez des RIB incomplets et/ou des montants de mauvais sens. Ces éléments ont été rejetés.')

    else if (lInCodeErr=CGE_PASERREUR) then
      PGIInfo('Les pièces et le ficher CFONB ont été correctement générés.', Ecran.Caption) ;
    end
  else if (lInCodeErr=CGE_PASERREUR)
    then PGIInfo('Les pièces comptables ont été correctement générées.', Ecran.Caption)
    else PGIInfo('Le traitement n''a pas pu aboutir.', Ecran.Caption) ;

end;

Initialization
  registerclasses ( [ TOF_CPGENERENCADECA ] ) ;
end.
