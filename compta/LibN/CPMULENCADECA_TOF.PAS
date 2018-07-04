{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/12/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULENCADECA ()
Mots clefs ... : TOF;CPMULENCADECA
*****************************************************************}
Unit CPMULENCADECA_TOF ;

Interface

Uses  StdCtrls,
      {$IFDEF VER150} Variants, {$ENDIF}
      Controls,
      Classes,
      Windows,
      {$IFDEF EAGLCLIENT}
        eMul,
        MaineAGL,
      {$ELSE}
        db,
        {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
        mul,
        FE_Main,
        hdb,
        dbGrids,
      {$ENDIF}
//      saisie,     // pour TrouveEtLanceSaisie
      SaisBor,    // Pour lanceSaisieFolio
      CPSaisiePiece_Tof,     // saisie paramètrable
      forms,
      sysutils,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      AGLInit,               // TheMulQ
      UTOF
      ,UTob
      ,Ent1            // pour ExoToDates
      ,ULibEncaDeca
      ,Saisutil         // pour QuelExo
      ,HQry
      ,HStatus          // pour InitMove et MoveCur et FiniMove
      ,ParamSoc         // pour le GetParamSoc
      ,TofVerifRib      // pour CPLanceFiche_VerifRib
      ,CPARAMGENER_TOM  // pour CPLanceFiche_ParamGener
      ,LettUtil         // rour DateCorrecte
      ,uLibWindows      // pour TraductionTHMultiValComboBox
      ;

Procedure CPLanceFiche_MulEncaDeca ( vStRange, vStLequel, vStArgs : String ) ;

Type
  TOF_CPMULENCADECA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    // Controles - Forme
  	Q               : THQuery ;
    {$IFDEF EAGLCLIENT}
      FListe        : THGrid ;
    {$ELSE}
      FListe        : THDBGrid ;
    {$ENDIF}
    Pages           : TPageControl ;
    FFiltres        : TComboBox ;
    HM              : THMsgBox ;
    BCtrlRib        : TButton ;

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
    
    {JP 04/06/07 : FQ 19747 : surcharge du changement de filtre}
    FLoadFiltre     : Boolean;
    FFiltreChange   : TNotifyEvent;

    // Variables
    Action        : String ;          // Stocke XXX de l'argument "ACTION=XXX"
    Flux          : String ;          // Code du flux paramétré (DEC ou ENC)
    CodeOPE       : String ;          // Code de l'opération paramétré (DEC ou ENC)
    Operation     : TTypeEncaDeca ;   // Stocke les infos relatives au fonctionnement de l'opération
    TOBScenario   : TOB ;             // Données du scénario sélectionné

    // EVT fiche scénario
    procedure FFluxChanged        (Sender: TObject) ;
    procedure FOperationChanged   (Sender: TObject);
    procedure FScenarioChanged    (Sender: TObject);
    procedure FAvecDateEcheClick  (Sender: TObject);

    // EVT fiche critères
    procedure MultiDossierChanged ( Sender : TObject ) ;
    procedure FTICChanged         (Sender: TObject);
    procedure E_EXERCICEChange    (Sender: TObject);
    procedure SpoolerChanged      (Sender: TObject);
    procedure FCategorieChanged   (Sender : TObject);
    procedure FSuiviTresoChanged  (Sender : TObject);
    procedure FTraiteEditeChanged (Sender: TObject);

    // EVT fiche boutons / liste
    procedure FormKeyDown         (Sender: TObject; var Key: Word; Shift: TShiftState) ;
    procedure BCtrlRibClick       (Sender: TObject);
    procedure BParamScenarioClick (Sender: TObject);
    procedure FListeDblClick      (Sender: TObject);
    procedure BOuvrirClick        (Sender: TObject);
    procedure SelectAllClick     (Sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);
    procedure ComboFiltreChange(Sender : TObject); {FQ 19747}
    {JP 07/06/07 : FQ 17001 : Affectation de la banque prévisionnelle}
    procedure AffecteBanqueClick(Sender : TObject);

    // Initialisations
    procedure AnalyseParametre( vStArg : String ) ;
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitMsgBox ;
    procedure InitCriteres ;
    procedure InitScenario ;
    procedure InitRepSpooler ;
    Procedure InitParPays ;
    procedure InitBanquePrevi; {JP 07/06/07 : FQ 17001}

    // Gestion scénario
    procedure ChargeOperation ;
    procedure ChargeScenario ;
    procedure majPanelScenario ;
    procedure majFiltreGeneral ;
    Procedure MajFiltreNaturePIece ; //XMG 05/04/04
    procedure CompleteTobScenario ;
    procedure GestionAccesScenario ;
    procedure GestionAccesDateEcheance ;
    procedure GestionAccesOngletEdition ;
    procedure GestionAccesOngletExport ;
    procedure GestionAccesSpooler ;
    procedure GestionAccesExport ;
    procedure GestionAccesOperation ;
    procedure GestionAccesNumChq ;
    procedure GestionAccesNumBor ;
    procedure GestionAccesGeneraux ;
    procedure GestionAccesEtab ;

    // Traitement chargement des échéances sélectionnées
    function  VerifieParametres : Boolean ;
    function  ConstitueOrigine : TOB ;
    procedure PrechargeOrigine ( vTobOrigine : TOB ) ;
    function  CleEcriture : String ;
    function  ReqEcriture : String ;
    function  Traitement( vTobOrigine : TOB ) : integer ;
    function  ConstitueOrigineMS : TOB ;
    Function  DatesOk : Boolean ;

    // Divers
    function  GetConditionSuiviTreso : String ;
    {Cette fonction renvoie la valeur du champ sChamp, que l'on travaille avec Q.FindField et
     GetField : cela évite de dupliquer les traitements entre NbSelected et AllSelected}
    function GetChamp(Q : THQuery; sChamp : string) : Variant;

    function  GetChampsOpti : string ;


  end ;

Implementation

uses
    {$IFDEF MODENT1}
    ULibExercice,
    CPTypeCons,
    CPProcMetier,
    {$ENDIF MODENT1}
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    ,UtilPGI, UlibEncaDecaESP, HPanel, Menus
    , uLibPieceCompta // TPieceCompta
    , Commun {JP 07/06/07 : FQ 15793 : SetPlusBanqueCp}
    , CPAFFECTEBQPREVI_TOF {JP 07/06/07 : FQ 17001}
    , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }


//==============================================================================
Procedure CPLanceFiche_MulEncaDeca ( vStRange, vStLequel, vStArgs : String ) ;
begin
  AGLLanceFiche('CP', 'CPMULENCADECA', vStRange, vStLequel, vStArgs) ;
end ;
//==============================================================================

//===================================
//======== EVENEMENTS TOF ===========
//===================================

procedure TOF_CPMULENCADECA.OnNew ;
begin
  Inherited ;
  // Placement de la 1ère page
  Pages.ActivePage        := TTabSheet(GetControl('PPARAMETRES',True)) ;
end ;

procedure TOF_CPMULENCADECA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMULENCADECA.OnUpdate ;
begin
  Inherited ;

end ;

procedure TOF_CPMULENCADECA.OnLoad ;
var lStWhere   : String ;
    lStCond    : String ;
    lStBidon   : String ; // chaine bidon pour passage arguments var
begin
  Inherited ;

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

  // condition Sur mode de paiement : Si tous les modes d'une catégorie précise, ajouter filtre sur la liste des modes
  if (FCategorie.ItemIndex > 0) and FModePaiement.Tous and ( FModePaiement.Values.Count > 0 ) then
    begin
    TraductionTHMultiValComboBox( FModePaiement, lStCond, lStBidon, 'E_MODEPAIE', False ) ;
    if lStWhere <> ''
       then lStWhere := lStWhere + ' AND ' + lStCond
       else lStWhere := lStCond ;
    end ;

  // condition sur le Suivi de trésorerie
  if lStWhere <> ''
    then lStWhere := lStWhere + ' AND ' + GetConditionSuiviTreso
    else lStWhere := GetConditionSuiviTreso ;


  // MAJ XX_WHERE
  SetControlText('XX_WHERE', lStWhere ) ;


end ;

procedure TOF_CPMULENCADECA.OnArgument (S : String ) ;
begin
  Inherited ;
  FLoadFiltre := False; {FQ 19747}

  // Analyse des paramètres
  AnalyseParametre( S ) ;

  //Initialization par Pays ;
  InitParPays ;

  // Initialisation de l'interface
  InitComposants ;
  InitEvenements ;
  InitMsgBox ;

  // Initialisation des onglets du scénario
  InitScenario ;

  // Initialisation des critères
  InitCriteres ;

  {JP 07/06/07 : FQ 17001 : gestion des banques prévisionnelles}
  InitBanquePrevi;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPMULENCADECA.OnClose ;
begin
  Inherited ;

  if Assigned(HM) then
    FreeAndNil(HM) ;
    
  if Assigned(TobScenario) then
    FreeAndNil( TOBScenario ) ;

end ;

procedure TOF_CPMULENCADECA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULENCADECA.OnCancel () ;
begin
  Inherited ;
end ;

//======================================
//======== EVENEMENTS FICHES ===========
//======================================


procedure TOF_CPMULENCADECA.FTICChanged(Sender: TObject);
begin
  if not FLoadFiltre then {FQ 19747}
    if (VH^.PaysLocalisation<>CodeISOES) or (GetControlEnabled('E_GENERAL')) then //XVI 24/02/2005
       SetControlText('E_GENERAL', '') ;
  GestionAccesGeneraux ;
end;

procedure TOF_CPMULENCADECA.FAvecDateEcheClick(Sender: TObject);
begin
  GestionAccesDateEcheance ;
end;

procedure TOF_CPMULENCADECA.FFluxChanged(Sender: TObject);
begin
  GestionAccesOperation ;
  GestionAccesScenario ;
end;

procedure TOF_CPMULENCADECA.FOperationChanged(Sender: TObject);
begin
  ChargeOperation ;
  GestionAccesScenario ;
end;

procedure TOF_CPMULENCADECA.FScenarioChanged(Sender: TObject);
begin
  ChargeScenario ;
end;

procedure TOF_CPMULENCADECA.BCtrlRibClick(Sender: TObject);
Var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(Pages) ;
  if (StWRib = '') then Exit;
  // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
  if ({$IFDEF EAGLCLIENT}
      // pas de tout sélectionner en CWAS
      {$ELSE}
      (Not FListe.AllSelected) and
      {$ENDIF}
      (FListe.NbSelected>0) and (FListe.NbSelected<100)) then
    begin
    // Rajoute une clause au WHERE
    StWRib := StWRib+' AND (';
    for i:=0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookmark(i) ;
      StWRib := StWRib +' (E_NUMEROPIECE='+ IntToStr( GetField('E_NUMEROPIECE') )
                       +' AND E_NUMLIGNE='+ IntToStr( GetField('E_NUMLIGNE') )
                       +' AND E_JOURNAL="'+ GetField('E_JOURNAL') +'") OR';
      end;
    // Efface le dernier OR et rajoute ')'
    StWRib := Copy( StWRib, 1, length(StWRib) - 2 );
    StWRib := StWRib +')';
    end;

  if (Operation.Code = 'LVN') or (Operation.Code = 'CVN')
    then CPLanceFiche_VerifRib('IBAN;WHERE='+StWRib)
    else CPLanceFiche_VerifRib('WHERE='+StWRib) ;

end;

procedure TOF_CPMULENCADECA.FListeDblClick(Sender: TObject);
var AA       : TActionFiche;
    sMode    : String ;
    lDossier : String ;
    lBoLocal : Boolean ;
    M        : RMVT ;
begin

  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

  {$IFDEF EAGLCLIENT}
    TheMulQ := TFMul(Ecran).Q.TQ;
    TheMulQ.Seek( FListe.Row - 1 );
  {$ELSE}
    TheMulQ := TFMul(Ecran).Q;
  {$ENDIF}

  if V_PGI.Sav
    then   AA       := taModif
    else   AA       := taConsult ;
  lDossier := V_PGI.SchemaName ;
  lBoLocal := True ;
  if ( GetControlText('MULTIDOSSIER')<>'' ) then// EstMultiSoc and ...
    begin
    lDossier := GetField('SYSDOSSIER') ;
    lBoLocal := ( lDossier <> V_PGI.SchemaName) ;
{    if ( lDossier <> V_PGI.SchemaName) and ( not V_PGI.Sav ) then
      begin
      lBoLocal := False ;
      AA     := taConsult ;
      end ;
}    end ;

  sMode := GetField('E_MODESAISIE') ;
  if ((sMode<>'') and (sMode<>'-')) then
    begin
    if lBoLocal then
      LanceSaisieFolio(TheMulQ,AA) ;
    end
  else
    begin
    // Données d'entête
    FillChar(M,Sizeof(M),#0) ;

    M.Jal           := GetField('E_JOURNAL') ;
    M.Exo           := GetField('E_EXERCICE') ;
    M.DateC         := GetField('E_DATECOMPTABLE') ;
    M.Num           := GetField('E_NUMEROPIECE') ;
    M.Nature        := GetField('E_NATUREPIECE') ;
    M.ModeSaisieJal := '-' ;
    M.Simul         := 'N' ;
    M.ANouveau      := False ;
    
    LanceSaisieParam( nil, AA, M, FALSE, lDossier ) ;
    end ;

end;

procedure TOF_CPMULENCADECA.BParamScenarioClick(Sender: TObject);
begin
  if V_PGI.Superviseur
    then CPLanceFiche_ParamGener('',FScenario.Value,'ACTION=MODIFICATION')
    else CPLanceFiche_ParamGener('',FScenario.Value,'ACTION=CONSULTATION') ;

  // Remise à jour des infos
  FScenario.Reload ;
  ChargeScenario ;

end;

procedure TOF_CPMULENCADECA.E_EXERCICEChange(Sender: TObject);
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

procedure TOF_CPMULENCADECA.SpoolerChanged(Sender: TObject);
begin
  GestionAccesSpooler ;
end;


//====================================
//======== INITIALISATIONS ===========
//====================================

procedure TOF_CPMULENCADECA.InitComposants;
begin

  // Controles - Formes
    Q             := THQuery(       GetControl('Q',         True) ) ;
  {$IFDEF EAGLCLIENT}
    FListe        := THGrid(        GetControl('FListe',    True) ) ;
  {$ELSE}
    FListe        := THDBGrid(      GetControl('FListe',    True) ) ;
  {$ENDIF}
  Pages           := TPageControl(  GetControl('Pages',     True) ) ;
  FFiltres        := TComboBox(     GetControl('FFILTRES',  True) ) ;
  BCtrlRib        := TButton(       GetControl('BCtrlRib',  True) );

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
  if Flux = 'ENC' then //XVI 24/02/2005
     Ecran.Caption := traduirememoire('Génération des encaissements')
  else
  if Flux = 'DEC' then
     Ecran.Caption := Traduirememoire('Génération des décaissements') ;
  UpdateCaption(Ecran) ;

  SetControlProperty('ETABLISSEMENTPAYEUR', 'TAG', -9979 ) ;

  {JP 30/10/07 : FQ 21432 : Glyph du bouton valider "VVVert"}
  TFMul(Ecran).BOuvrir.GlobalIndexImage := 'Z0184_S16G1';
end;

procedure TOF_CPMULENCADECA.InitCriteres;
begin

  // Param onglet tables libres
  LibellesTableLibre( TTabSheet(GetControl('PLibres', True)), 'TT_TABLE', 'T_TABLE', 'T') ;

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
  // Titre de la fenêtre
  if VH^.PaysLocalisation<>CodeISOES then
  begin
     if Flux = 'ENC'
       then SetControlText('E_NATUREPIECE', 'AC;FC;')
       else if Flux = 'DEC'
            then SetControlText('E_NATUREPIECE', 'AF;FF;')
            else SetControlText('E_NATUREPIECE', 'AC;AF;FC;FF;');
  End ; //XVI 24/02/2005 //XMG Pour l'espagne, c'est par rapport à l'action.....

  // Journal..(<<Tous>> par défaut)
  SetControlProperty('E_JOURNAL', 'ITEMINDEX', 0) ;

  // Suivi tréso (MIX par défaut)
  SetControlText('SUIVITRESO', 'X') ;

  GestionAccesGeneraux ;

  // Cacher l'onglet contenant les champs XX_WHERE "pinvisible"
  TTabSheet(GetControl('PINVISIBLE',True)).TabVisible := False ;

end;

procedure TOF_CPMULENCADECA.InitEvenements;
begin

  // Ecran
  if Assigned( Ecran.OnKeyDown )
    then FOldKeyDown := Ecran.OnKeyDown
    else FOldKeyDown := nil ;
  Ecran.OnKeyDown         := FormKeyDown ;

  // liste
  FListe.OnDblClick       := FListeDblClick ;

  // Controles - Onglet scénario
  FFlux.OnChange          := FFluxChanged ;
  FOperation.OnChange     := FOperationChanged ;
  FScenario.OnChange      := FScenarioChanged ;
  FAvecDateEche.OnClick   := FAvecDateEcheClick ;
  BParamScenario.OnClick  := BParamScenarioClick ;

  // onglet Edition
  TCheckBox(GetControl('Spooler', True)).OnClick := SpoolerChanged ;

  // Controles - Onglet standard
  FTIC.OnClick            := FTICChanged ;
  FCategorie.OnChange     := FCategorieChanged ;
  FSuiviTreso.OnClick     := FSuiviTresoChanged ;
  FTraiteEdite.OnClick    := FTraiteEditeChanged ;
  ThValComboBox( GetControl('MultiDossier', True) ).OnChange := MultiDossierChanged ;

  // Onglet Ecritures
  THValComboBox(GetControl('E_EXERCICE',  True)).OnChange := E_EXERCICEChange ;

  // boutons...
  BCtrlRib.OnClick                             := BCtrlRibClick ;
  TButton(GetControl('BOuvrir', True)).OnClick := BOuvrirClick ;

  TButton(GetControl('BSELECTALL',True)).OnClick := SelectAllClick;
  
  {JP 04/06/07 : FQ 19747 : Ajout d'un booléen pour savoir si le filtre est en chargement}
  FFiltreChange := FFiltres.OnChange;
  FFiltres.OnChange := ComboFiltreChange;
end;

procedure TOF_CPMULENCADECA.InitScenario;
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

  // Initialisation des paramètres d'exports
  SetControlText('REMISEREF',         '' );
  SetControlText('REMISEDATE',        DateToStr(V_PGI.DateEntree) );
  SetControlText('FRAISIMPUTATION',   '14' );

  // Initialisation de l'établissement payeur
  lStEtab := GetParamSocSecur('SO_ETABLISDEFAUT', '')  ;
  SetControlText('ETABLISSEMENTPAYEUR', lStEtab ) ;

end;

procedure TOF_CPMULENCADECA.InitRepSpooler;
begin
  SetControlText( 'RepSpooler', CheminSpoolerPourEncaDeca( Operation ) ) ;
end;

//=============================================
//======== GESTION ACCES AUC CHAMPS ===========
//=============================================

procedure TOF_CPMULENCADECA.GestionAccesDateEcheance;
begin
  FDateEche.Enabled := FAvecDateEche.Checked ;
end;

procedure TOF_CPMULENCADECA.GestionAccesScenario ;
begin
  if Operation.Code = 'TOUS'
    then FScenario.Plus := ' CPG_FLUXENCADECA="' + FFlux.Value + '"'
    else FScenario.Plus := ' CPG_FLUXENCADECA="' + FFlux.Value
                           + '" AND CPG_TYPEENCADECA="' + Operation.Code + '"' ;

  if GetControlText('MULTIDOSSIER')<>'' then
    FScenario.Plus := FScenario.Plus + ' AND CPG_MULTISOC="X" ' ;

  FScenario.ItemIndex := 0 ;
  ChargeScenario ;
  If VH^.PaysLocalisation=CodeISOES then
     MajFiltreNaturePiece ; //XVI 24/02/2005
end;

procedure TOF_CPMULENCADECA.GestionAccesOngletEdition ;
Var lBoTestEtat    : Boolean ;
begin
  // Onglet edition / export visible ?
  TTabSheet(GetControl('PEDITION')).TabVisible := ( Operation.Code <> '' ) and ( Operation.accesEdition <> taJamais ) ;

  // Accès aperçu / spooler si scenario prévoit une édition
  lBoTestEtat    := ( TobScenario <> nil )  and ( TobScenario.GetValue('CPG_EDITE') = 'X' ) ;
  SetControlEnabled('APERCU',      lBoTestEtat ) ;
  SetControlEnabled('SPOOLER',     lBoTestEtat ) ;
  if not lBoTestEtat then
    begin
    SetControlText('APERCU', '-') ;
    SetControlText('SPOOLER', '-') ;
    end ;

  GestionAccesSpooler ;
end;

procedure TOF_CPMULENCADECA.GestionAccesSpooler;
Var lBoTestSpooler : Boolean ;
    lBoTestEtat    : Boolean ;
begin
  lBoTestEtat    := ( TobScenario <> nil )  and ( TobScenario.GetValue('CPG_EDITE') = 'X' ) ;
  lBoTestSpooler := ( GetCheckBoxState('Spooler') = cbChecked ) ;

  SetControlEnabled('RepSpooler',       lBoTestSpooler) ;
  SetControlEnabled('TRepSpooler',      lBoTestSpooler) ;
  SetControlEnabled('XFichierSpooler',  lBoTestSpooler) ;

  SetControlEnabled('Apercu',           lBoTestEtat ) ; //and (not lBoTestSpooler) ) ; // FQ 17883 : Accès au spooler sans aperçu

  If lBoTestSpooler then
     begin
     if GetControlText('RepSpooler') = ''
        then InitRepSpooler ;
//     SetControlText('Apercu', 'X' ) ; // FQ 17883 : Accès au spooler sans aperçu
     end ;

end;

//================================
//======== CHARGEMENTS ===========
//================================

procedure TOF_CPMULENCADECA.ChargeOperation;
begin
  if FOperation.ItemIndex = 0
    then Operation := ChargeTypeEncaDeca( 'TOUS' )
    else Operation := ChargeTypeEncaDeca( FOperation.Value ) ;
end;


procedure TOF_CPMULENCADECA.ChargeScenario;
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
  // Affichage zone de saisie du numéro de chèque
  GestionAccesNumChq ;
  // Affichage zone de saisie du numéro de bordereau
  GestionAccesNumBor ;
  // Accès aux paramètrages de l'édition
  GestionAccesOngletEdition ;
  // Accès aux paramètrages de l'export
  GestionAccesOngletExport ;

  if VH^.PaysLocalisation=CodeISOEs then //XVI 24/02/2005
     majFiltreGeneral ;
end;

//===================================
//======== MAJ DES CHAMPS ===========
//===================================

procedure TOF_CPMULENCADECA.majPanelScenario;
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

procedure TOF_CPMULENCADECA.AnalyseParametre(vStArg: String);
var lStChaine : String ;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt ( vStArg ) ;

    // Stocke le XXX de "ACTION=XXX" dans la variable Action
    if Pos( 'ACTION=' , Uppercase( lStChaine ) ) > 0 then
      begin
      Action := Copy(lStChaine, 8, length( lStChaine ) )
      end

    // Flux fixé par appel
    else if Pos( 'FLUX=' , Uppercase( lStChaine ) ) > 0 then
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

procedure TOF_CPMULENCADECA.BOuvrirClick(Sender: TObject);
var lTOBOrigine    : TOB ;             // Données sur les pièces d'origine
    lInCodeErr     : integer ;
begin

  // Maj des champs supplémentaires de paramètrage
  CompleteTobScenario ;

  // Test d'usage
  if not VerifieParametres then Exit ;

  // Creation de la Tob contenant les échéances d'origines
  if GetControlText('MULTIDOSSIER')<>''
    then lTOBOrigine := ConstitueOrigineMS
    else lTOBOrigine := ConstitueOrigine ;

  lInCodeErr := CGE_ERRTRAITEMENT ;
  if lTOBOrigine.Detail.Count > 0 then
     lInCodeErr := Traitement( lTOBOrigine ) ;

  // Codes erreur possibles :
  Case lInCodeErr of

    // ==> TRAITEMENT OK
    CGE_PASERREUR  : begin
                     PGIInfo('Le traitement s''est correctement effectué.', Ecran.Caption ) ;
                     TFMul(Ecran).BChercheClick( nil ) ;
                     end ;

    // ==> ANNULATION UTILISATEUR
    CGE_ANNULATION : PgiInfo('Traitement annulé par l''utilisateur.', Ecran.Caption ) ;

    // ==> Message d'erreur traité dans le traitement
    CGE_NOEXODOSSIER :  ; // msg déjà affiché dans traitement
    CGE_EXOCLOS      :  ; // msg déjà affiché dans traitement

    // ==> Message d'erreur précis
    CGE_ERRVALIDEPIECEMS    : PgiInfo('Traitement annulé suite à une erreur lors de la validation du règlement multi-société généré.', Ecran.Caption ) ;
    CGE_ERRVALIDEPIECE      : PgiInfo('Traitement annulé suite à une erreur lors de la validation du règlement généré.', Ecran.Caption ) ;
    CGE_ERRVALIDEESC        : PgiInfo('Traitement annulé suite à une erreur lors de la validation de l''escompte générée.', Ecran.Caption ) ;
    CGE_ERRNUMPIECE         : PgiInfo('Traitement annulé suite à une erreur lors de l''affectation des numéros de pièces.', Ecran.Caption ) ;
    CGE_ERRSAVEPIECE        : PgiInfo('Traitement annulé suite à une erreur lors de l''enregistrement des pièces.', Ecran.Caption ) ;
    CGE_ERRLETTPIECE        : PgiInfo('Traitement annulé suite à une erreur lors du lettrage des pièces.', Ecran.Caption ) ;
    CGE_ERRUPDATEECHE       : PgiInfo('Traitement annulé suite à une erreur lors de la mise à jour des échéances d''origine.', Ecran.Caption ) ;
    CGE_ERREDITION          : PgiInfo('Traitement effectué mais un erreur est surnevue lors de l''édition du règlement.', Ecran.Caption ) ;
    CGE_ERREDITBOR          : PgiInfo('Traitement effectué mais un erreur est surnevue lors de l''édition du BOR.', Ecran.Caption ) ;

    // ==> Message d'erreur non déterminé
    CGE_ERRTRAITEMENT       : PGIInfo('Le traitement a échoué.', Ecran.Caption) ;
    CGE_ERRTRANSAC          : PGIInfo('Le traitement a échoué.', Ecran.Caption) ;
    else                      PGIInfo('Le traitement a échoué.', Ecran.Caption) ;

  end ;

  // Libération de la TOB
  if Assigned( lTobOrigine ) then
    begin
    lTOBOrigine.ClearDetail ;
    FreeAndNil( lTOBOrigine ) ;
    end;

end;

function TOF_CPMULENCADECA.VerifieParametres: Boolean;
begin
  Result := False ;

// Le résultat de la recherche est-il vide ?
{$IFDEF EAGLCLIENT}
	if (Q.TQ.Detail.Count = 0) then
{$ELSE}
	if (Q.Eof and Q.Bof) then
{$ENDIF}
    begin
    PGIInfo('Le résultat de la recherche est vide.', Ecran.Caption) ;
    exit ;
    end ;

// Au moins une ligne de sélectionnée en mode normale, (pas de test en mode inverse pour l'instant)
	if ( FListe.NbSelected = 0 ) and ( not FListe.AllSelected ) then
    begin
    PGIInfo('Aucune ligne à traiter. Vous devez sélectionner au moins une échéance.',Ecran.Caption) ;
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

// Test N° de chèque
  if ( TobScenario.GetValue('CPG_AVECNUMCHEQUE')='X' ) and
     ( ( not IsNumeric( TobScenario.GetValue('NUMCHEQUE') ) ) or
       ( StrToInt( TobScenario.GetValue('NUMCHEQUE') ) <= 0 )
     ) then
    begin
    PGIInfo('Le numéro de chèque saisi n''est pas valide. Vous devez saisir un entier supérieur à zéro.', Ecran.Caption) ;
    SetActiveTabSheet('PEDITION') ;
    SetFocusControl('NUMCHEQUE') ;
    exit ;
    end;

// Test journal de substitution pour les pièces sans contrepartie
  if ( TobScenario.GetString('CPG_MULTIETAB') = 'X' ) and ( GetParamSocSecur( 'SO_CPJALENCADECA', '' ) = '' ) then
    begin
    PGIInfo('Vous devez renseigner le journal de substitution (paramètre sociétés - suivi-tiers) pour pouvoir utiliser un scénario multi-établissement.', Ecran.Caption) ;
    exit ;
    end ;

// Test journal d'escompte
  if ( TobScenario.GetValue('CPG_EscMethode')<>'RIE') and ( GetParamSocSecur('SO_CPJALESCOMPTE', '') = '' ) then
    PGIInfo ( 'Le journal d''escompte n''est pas renseigné dans les paramètres société. Aucune escompte ne pourra être générée.'
             , Ecran.Caption ) ;

// Test paramètres d'export
{
  if ( TobScenario.GetValue('CPG_CFONBEXPORT')='X' ) then
    begin
    // Test fichier
    if Trim(TobScenario.GetValue('EXPORTFICHIER')) = ''  then
      begin
      PGIInfo('Le scénario sélectionné comprend un export CFONB. Vous devez saisir un nom de fichier valide.', Ecran.Caption) ;
      SetActiveTabSheet('PEXPORT') ;
      SetFocusControl('EXPORTFICHIER') ;
      exit ;
      end;
    if (VH^.PaysLocalisation=CodeISOES) then  //XVI 24/02/2005
    begin
       if ( TobScenario.GetValue('CPG_CFONBFORMAT')='N32' ) and
          ( Trim(GetControlText('IDCEDENTE'))='' )              then
       Begin
          PGIInfo('Le scénario sélectionné comprend un export CFONB au format Norme 32. Vous devez saisir l''identifiant du cédant.', Ecran.Caption) ;
          SetActiveTabSheet('PEXPORT') ;
          SetFocusControl('REMISEREF') ;
          exit ;
       end ;
    end else
    // Test date de remise
    if Trim(GetControlText('REMISEREF'))='' then
      begin
      PGIInfo('Le scénario sélectionné comprend un export CFONB. Vous devez saisir la référence de la remise.', Ecran.Caption) ;
      SetActiveTabSheet('PEXPORT') ;
      SetFocusControl('REMISEREF') ;
      exit ;
      end;
    // Test référence de remise
    if not IsValidDate( GetControlText('REMISEDATE') ) then
      begin
      PGIInfo('Le scénario sélectionné comprend un export CFONB. Vous devez saisir la date de la remise.', Ecran.Caption) ;
      SetActiveTabSheet('PEXPORT') ;
      SetFocusControl('REMISEDATE') ;
      exit ;
      end;
    end ;
}
    // Message d'avertissement pour l'escompte
//  if AvecEscompte( TobScenario ) then
//    if HM.Execute(9,gszCaption,'')<>mrYes then Exit ;

(*
// Message avertissement pour sélection inversée
	if swapSelect then
  	if PGIAsk('Vous êtes en mode "sélection inversée", Le traitement peut être long, voulez-vous continuez ?',gszCaption) <> mrYes
    then  exit;
*)

// Confirmation du traitement
  if PGIAsk(Format(traduireMemoire('Confirmez-vous l''application du scénario %s - %s sur les échéances sélectionnées ?'),
                   [TobScenario.GetValue('CPG_CODE'), TobScenario.GetValue('CPG_LIBELLE')])) <> mrYes                     then //XVI 24/02/2005
     Exit ;

  // Tout est ok
  Result := True ;

end;

procedure TOF_CPMULENCADECA.InitMsgBox;
begin
  // Init HM
  HM:= THMsgBox.create(FMenuG);
  HM.Mess.Clear ;
  HM.Mess.Add('0;?Caption?;Vous devez renseigner la zone "Compte de sélection";W;O;O;O;');
  HM.Mess.Add('1;?Caption?;Les lignes d''écritures ont des modes de paiement et des échéances qui diffèrent. Confirmez-vous le traitement ?;E;YN;N;N;');
  HM.Mess.Add('2;?Caption?;Les lignes d''écritures ont des modes de paiement qui diffèrent. Confirmez-vous le traitement ?;E;YN;N;N;');
  HM.Mess.Add('3;?Caption?;Les lignes d''écritures ont des echéances qui diffèrent. Confirmez-vous le traitement ?;E;YN;N;N;');
  HM.Mess.Add('4;?Caption?;Le compte de génération ne doit pas être identique au compte des échéances sélectionnées;W;O;O;O;');
// En attente modif lettrage SBO :
//  HM.Mess.Add('4;?Caption?;Le compte de génération est identique au compte des échéances sélectionnées. Confirmez-vous le traitement ?;Q;YN;N;N;');
  HM.Mess.Add('5;?Caption?;Certains comptes de banque n''ont aucun journal associé;W;O;O;O;');
  HM.Mess.Add('6;?Caption?;Le traitement sur ces comptes n''a pas été effectué;W;O;O;O;');
  HM.Mess.Add('7;?Caption?;Les lignes d''écritures ont des modes de paiement qui diffèrent. Veuillez saisir le Mode de paiement de contrepartie.;W;O;O;O;');
  HM.Mess.Add('8;?Caption?;Les lignes d''écritures ont des modes de paiement et des échéances qui diffèrent. Vous devez renseigner un mode de paiement;W;O;O;O;');
  HM.Mess.Add('9;?Caption?;Vous avez demandez une génération avec escompte. Confirmez-vous le traitement ?;Q;YN;N;N;');
  HM.Mess.Add('10;?Caption?;Certaines lignes ont  une échéance antérieure à la date du jour (Pas d''escompte). Confirmez-vous le traitement ?;Q;YN;N;N;');

  HM.Mess.Add('11;?caption?;La date d''échéance doit respecter la plage de saisie autorisée;W;O;O;O;') ;
  HM.Mess.Add('12;?caption?;Vous devez renseigner une date valide.;W;O;O;O;') ;
  HM.Mess.Add('13;?caption?;La date de règlement que vous avez renseignée est sur un exercice non ouvert.;W;O;O;O;') ;
  HM.Mess.Add('14;?caption?;La date de règlement que vous avez renseignée est sur un exercice non ouvert.;W;O;O;O;') ;
  HM.Mess.Add('15;?caption?;La date de règlement que vous avez renseignée est antérieure à la clôture provisoire.;W;O;O;O;') ;
  HM.Mess.Add('16;?caption?;La date de règlement que vous avez renseignée est antérieure à la clôture définitive.;W;O;O;O;') ;

  HM.Mess.Add('17;?caption?;La date que vous avez renseignée n''est pas valide;W;O;O;O;') ;
  HM.Mess.Add('18;?caption?;La date que vous avez renseignée n''est pas dans un exercice ouvert;W;O;O;O;') ;
  HM.Mess.Add('19;?caption?;La date que vous avez renseignée est antérieure à une clôture;W;O;O;O;') ;
  HM.Mess.Add('20;?caption?;La date que vous avez renseignée est antérieure à une clôture;W;O;O;O;') ;
  HM.Mess.Add('21;?caption?;La date que vous avez renseignée est en dehors des limites autorisées;W;O;O;O;') ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 06/04/2004
Modifié le ... : 23/04/2004
Description .. : Construit une TOB contenant les ecritures sélectionnées
Suite ........ : dans le mutlicritères.
Suite ........ : 
Suite ........ : Modif 06/04/2004 : Les écritures sont maintenant divisées
Suite ........ : en 2 catégories : les échéances client et les échéances
Suite ........ : fournisseurs...
Suite ........ : 
Suite ........ : Modif 23/04/2004 : Annulation de la modif ci-dessus...
Mots clefs ... :
*****************************************************************}
function TOF_CPMULENCADECA.ConstitueOrigine : TOB ;
Var lTobEcr    : Tob ;
    i          : Integer ;
    lQEcr      : TQuery ;
begin

  // Libération si besoin
  Result        := TOB.Create( 'TOB_ORIGINE', nil, -1 ) ;

  // Chargement
  if FListe.AllSelected

    // On charge toute la liste
    then PrechargeOrigine( Result )

    // Uniquement la sélection
    else if FListe.nbSelected > 0 then
      begin
      for i:=0 to FListe.NbSelected-1 do
        begin
        FListe.GotoLeBookmark(i) ;
        lTobEcr := TOB.Create('ECRITURE', Result, -1 ) ;
        lQEcr := OpenSQL( ReqEcriture + CleEcriture , True ) ;
        lTobEcr.SelectDB( '' , lQEcr ) ;
        Ferme( lQEcr ) ;
        end ;
      end ;

end;

procedure TOF_CPMULENCADECA.PrechargeOrigine ( vTobOrigine : TOB ) ;
Var //lQEcr  : TQuery ;
    lStReq : String ;
begin

  // FQ22173 : Gestion des gros volumes : optimisation des chargements des infos tiers
  lStReq := ReqEcriture + RecupWhereCritere(Pages) ;
  vTobOrigine.LoadDetailFromSQL(lStReq ) ;
{
  lQEcr := OpenSelect( stReq , vDossier ) ;
  if not lQEcr.Eof then
    vTobOrigine.LoadDetailDB( 'ECRITURE', '', '', lQEcr, True, True ) ;
  Ferme(lQEcr) ;
}

end;

function TOF_CPMULENCADECA.CleEcriture: String;
Var lStJal      : String ;
    lStExo      : String ;
    lStDate     : TDateTime ;
    lInNumP     : Integer ;
    lInNumL     : Integer ;
    lInNumEche  : Integer ;
begin
  lStJal      := GetField('E_JOURNAL') ;
  lStExo      := GetField('E_EXERCICE') ;
  lStDate     := GetField('E_DATECOMPTABLE') ;
  lInNumP     := GetField('E_NUMEROPIECE') ;
  lInNumL     := GetField('E_NUMLIGNE') ;
  lInNumEche  := GetField('E_NUMECHE') ;

  Result := ' WHERE E_JOURNAL="'        + lStJal                 + '" AND '
                 + 'E_EXERCICE="'       + lStExo                 + '" AND '
                 + 'E_DATECOMPTABLE="'  + UsDateTime(lStDate)    + '" AND '
                 + 'E_NUMEROPIECE='     + IntToStr(lInNumP)      + ' AND '
                 + 'E_NUMLIGNE='        + IntToStr(lInNumL)      + ' AND '
                 + 'E_NUMECHE='         + IntToStr(lInNumEche) ;
end;

function TOF_CPMULENCADECA.ReqEcriture: String;
begin
{
  result := 'SELECT ECRITURE.*, T_NATUREAUXI NATUREAUXI, G_NATUREGENE NATUREGENE' ;
  result := result + ' FROM ECRITURE'
                   + ' LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE'
                   + ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' ;
}
  // FQ22173 : Gestion des gros volumes : optimisation des chargements des infos tiers
  result := 'SELECT ' + GetChampsOpti + ', ' + TPieceCompta.GetChampsAuxi + ', ' + TPieceCompta.GetChampsGene
               + ' FROM ECRITURE'
                   + ' LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE'
                   + ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' ;
end;

procedure TOF_CPMULENCADECA.CompleteTobScenario;
var lStTmp    : String ;
    QQ        : TQuery ;
    idxFrais  : Integer ;
begin
  if TobScenario = nil then Exit ;

  // Paramètres contextuelles
  TobScenario.PutValue('DATECOMPTABLE',   StrToDate( FDateComptable.Text ) ) ;
  if FAvecDateEche.Checked
    then TobScenario.PutValue('FORCERDATEECHE', 'X' )
    else TobScenario.PutValue('FORCERDATEECHE', '-' ) ;
  TobScenario.PutValue('DATEECHE',        StrToDate( FDateEche.Text ) ) ;
  TobScenario.PutValue('ETABLISSEMENTPAYEUR',  GetControlText('ETABLISSEMENTPAYEUR') ) ;

  // Paramètres d'éditions
  if GetCheckBoxState('APERCU') = cbChecked
    then TobScenario.PutValue('APERCU', 'X' )
    else TobScenario.PutValue('APERCU', '-' ) ;
  if GetCheckBoxState('SPOOLER') = cbChecked
    then TobScenario.PutValue('SPOOLER', 'X' )
    else TobScenario.PutValue('SPOOLER', '-' ) ;
  if GetCheckBoxState('XFICHIERSPOOLER') = cbChecked
    then TobScenario.PutValue('XFICHIERSPOOLER', 'X' )
    else TobScenario.PutValue('XFICHIERSPOOLER', '-' ) ;
  TobScenario.PutValue('REPSPOOLER',  GetControlText('REPSPOOLER') ) ;

  // Paramètres d'export
  TobScenario.PutValue('EXPORTFICHIER',    GetControlText('EXPORTFICHIER') ) ;
  if GetCheckBoxState('EXPORTAPERCU') = cbChecked
    then TobScenario.PutValue('EXPORTAPERCU', 'X' )
    else TobScenario.PutValue('EXPORTAPERCU', '-' ) ;
  lStTmp := GetControlText('REMISEREF') ;
  TobScenario.PutValue('REMISEREF',        GetControlText('REMISEREF') ) ;
  if VH^.PaysLocalisation=CodeISOES then
     TobScenario.PutValue('REMISEDATE',       StrToDate(GetControlText('REMISEDATE')))
  else
  begin
     lStTmp := FormatDateTime('yyyymmdd', StrToDate(GetControlText('REMISEDATE')) ) ;
     TobScenario.PutValue('REMISEDATE',       FormatDateTime('yyyymmdd', StrToDate(GetControlText('REMISEDATE')) ) ) ;
  End ; //XVI 24/02/2005
  lStTmp := GetControlText('FRAISIMPUTATION') ;
  TobScenario.PutValue('FRAISIMPUTATION',  GetControlText('FRAISIMPUTATION') ) ;
  if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
  begin
     TobScenario.PutValue('IDCEDENTE',  GetControlText('IDCEDENTE') ) ;
     TobScenario.PutValue('EFFETSPHYS',  GetControlText('EFFETSPHYS') ) ;
  End ;

  // Numéro de chèque intial : 3 cas possible Chèque, BOR, Traite
  if TobScenario.GetValue('CPG_AVECNUMCHEQUE')='X' then
    begin
    TobScenario.PutValue('NUMCHEQUE', GetControlText('NUMCHEQUE') ) ;
    TobScenario.PutValue('REFCHEQUE', GetControlText('REFCHEQUE') ) ;
    end ;
  // Numéro de bordereau
  if ( TobScenario.GetValue('CPG_BORDEREAUEXP')='X' ) then
    TobScenario.PutValue('NUMENCADECA', GetControlText('NUMBORDEREAU') ) ;

  // Sélection uniquement de Tic/Tid
  if GetCheckBoxState('FTIC') = cbChecked
    then TobScenario.PutValue('TICTID', 'X' )
    else TobScenario.PutValue('TICTID', '-' ) ;

  //Paramètres spécifiques ESP
  if VH^.PaysLocalisation=CodeISOES then
     Begin
     if TobScenario.GetValue('CPG_CONDREMISE')<>'' then
        Begin
        TobScenario.ClearDetail ;
        with TOB.Create('CCONDREMISE', TobScenario, -1) do
          Begin
          PutValue( 'CCB_CONDREMISE' , TobScenario.GetValue('CPG_CONDREMISE')) ;
          if LoadDB then
             Begin
             QQ:=OpenSQL('select * from CFRAISREMISE '+
                         'where CFR_CONDREMISE="'+GetValue('CCB_CONDREMISE')+'" and '+
                               'CFR_TYPEFRAIS="'+TypeEncaDecaVersTypeFrais(TobScenario.GetValue('CPG_TYPEENCADECA'))+'"'+
                         'Order by CFR_FRAISREMISE' ,TRUE) ;
             LoadDetailDB('CFRAISREMISE','','',QQ,FALSE) ;
             Ferme(QQ) ;
             For IdxFrais:=0 to TobScenario.Detail[0].Detail.Count-1 do
               Begin
               if TobScenario.Detail[0].Detail[IdxFrais].GetValue('CFR_TVA')<>'' then
                  Begin
                  TobScenario.Detail[0].Detail[IdxFrais].AddChampSup('CFR_GENERALTVA',FALSE) ;
                  TobScenario.Detail[0].Detail[IdxFrais].PutValue('CFR_GENERALTVA',TVA2CPTE(VH^.RegimeDefaut,TobScenario.Detail[0].Detail[IdxFrais].GetValue('CFR_TVA'),TRUE)) ;
                  End ;
               End ;
             End ;
          End ;
        End ;
     End ; //XVI 24/02/2005
end;

procedure TOF_CPMULENCADECA.FCategorieChanged(Sender : TObject);
begin

  // Le 28/01/2003 BPY correction des bug de la fiche 11794
  // SBO 29/06/2005 : FQ 17774
  BCtrlRib.visible := (FCategorie.Value <> 'TRI') or(VH^.PaysLocalisation=CodeISOES) ; //XVI 24/02/2005

  //  remplissage du combo des mode de paiement
  if FCategorie.ItemIndex = 0
     then CategorieVersModePaiement( '' , FModePaiement )
     else CategorieVersModePaiement( FCategorie.Value , FModePaiement ) ;

end;


procedure TOF_CPMULENCADECA.InitParPays;
begin
  if VH^.PaysLocalisation=CodeISOES then
  Begin
     with THPanel(GetControl('PNL_FRA')) do
     Begin
       Visible:=FALSE ;
       SetControlProperty('PNL_ESP','top',top) ;
       SetControlProperty('PNL_ESP','left',Left) ;
       SetControlProperty('PNL_ESP','Height',Height) ;
       SetControlProperty('PNL_ESP','Width',Width) ;
     End ;
     SetControlVisible('PNL_ESP',TRUE) ;
     SetControlProperty('E_GENERAL','Tag',-9980) ;
     SetControlProperty('E_NATUREPIECE','Tag',-9980) ;
     SetControlVisible('ESCMETHODE',   FALSE ) ;
     SetControlVisible('TESCMETHODE',  FALSE ) ;
     SetControlVisible('ESCTAUX',   FALSE ) ;
     SetControlVisible('TESCTAUX',  FALSE ) ;
  End
  else
  if assigned(getcontrol('PNL_ESP')) then
    SetControlVisible('PNL_ESP',FALSE) ;

  // FQ17000 : Pb perso de liste
  if Flux = 'ENC'
    then TFMul(Ecran).SetDBListe( 'CPMULENCA' )
  else if Flux = 'DEC'
    then TFMul(Ecran).SetDBListe( 'CPMULDECA' ) ;

end;

procedure TOF_CPMULENCADECA.GestionAccesOperation;
Begin
  //on ne prend que les type d'enca/Déca que corresponent....
  SetControlProperty('OPERATION','PLUS',ConditionTypeEncaDeca(FFlux.Value)) ;
  if VH^.PaysLocalisation=CodeISOES then //XVI 24/02/2005
  Begin
    FOperation.ItemIndex:=0 ;
    GestionAccesScenario ;
  End ;
end;

procedure TOF_CPMULENCADECA.GestionAccesNumChq;
var lBoAvecNumChq : Boolean ;
begin
  lBoAvecNumChq := ( TobScenario <> nil ) and ( TobScenario.GetValue('CPG_AVECNUMCHEQUE')='X' ) ;
  SetControlVisible('TNUMCHEQUE', lBoAvecNumChq ) ;
  SetControlVisible('REFCHEQUE',  lBoAvecNumChq ) ;
  SetControlVisible('NUMCHEQUE',  lBoAvecNumChq ) ;
end;

procedure TOF_CPMULENCADECA.GestionAccesGeneraux;
begin
  SetControlEnabled('E_AUXILIAIRE',  not FTIC.Checked ) ;
  SetControlEnabled('TE_AUXILIAIRE', not FTIC.Checked ) ;
  if FTIC.Checked then
    begin
    SetControlText('E_AUXILIAIRE', '' ) ;
    SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGTIDTIC') ;
    end
  else SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGCOLLECTIF') ;
  SetControlProperty('E_GENERAL', 'PLUS', ' AND ' + GetConditionSuiviTreso )
end;

procedure TOF_CPMULENCADECA.FSuiviTresoChanged(Sender: TObject);
begin
  if not FLoadFiltre then {FQ 19747}
    if (VH^.PaysLocalisation<>CodeISOES) or (GetControlEnabled('E_GENERAL')) then //XVI 24/02/2005
      SetControlText('E_GENERAL', '') ;
  GestionAccesGeneraux ;
end;

procedure TOF_CPMULENCADECA.majFiltreGeneral;
var GeneralSel : String ;
begin
  if VH^.PaysLocalisation=CodeISOES then
     Begin
     GeneralSel:=TobScenario.GetValue('CPG_GENERALSEL') ;
     SetControltext('E_GENERAL',GeneralSel) ;
     SetControlEnabled('E_GENERAL',(trim(GeneralSel)='')) ;
     End ;
end;

Procedure TOF_CPMULENCADECA.MajFiltreNaturePIece ;
var StnaturePiece : String ;
Begin
  // Nature de pièce... (par défaut Opèraitons Diverses)
  StNaturePiece:='OD;' ;
  //Si Opèration Mise en portefauille alors (Factures et Avois) ;
  if FOperation.Value = 'MPF'
     then StNaturePiece:= 'AC;AF;FC;FF;' ;

  SetControlText('E_NATUREPIECE', StNaturePiece);
End ;

procedure TOF_CPMULENCADECA.FTraiteEditeChanged(Sender: TObject);
begin
  SetControlEnabled('E_NUMTRAITECHQ', FTraiteEdite.State = cbChecked ) ;
  SetControlEnabled('E_NUMTRAITECHQ_', FTraiteEdite.State = cbChecked ) ;
  SetControlEnabled('TE_NUMTRAITECHQ_', FTraiteEdite.State = cbChecked ) ;
end;

function TOF_CPMULENCADECA.Traitement( vTobOrigine : TOB ): integer;
begin

  if VerifieCoherenceSelection( vTobOrigine, TobScenario ) then
    begin
    // Traitement
    Result := ExecuteEncaDeca( vTobOrigine , TobScenario ) ;
    end
  else
    Result := CGE_ANNULATION ;

end;

procedure TOF_CPMULENCADECA.GestionAccesExport;
Var lBoExportCFONB : Boolean ;
begin

  // plus utilisé en mode FR
  if VH^.PaysLocalisation<>CodeISOES then Exit ;

  lBoExportCFONB := ( TobScenario <> nil )  and ( TobScenario.GetValue('CPG_CFONBEXPORT') = 'X' ) ;

  SetControlEnabled('EXPORTFICHIER',    lBoExportCFONB) ;
  SetControlEnabled('TEXPORTFICHIER',   lBoExportCFONB) ;
  SetControlEnabled('EXPORTAPERCU',     lBoExportCFONB) ;

  SetControlEnabled('REMISEREF',        lBoExportCFONB) ;
  SetControlEnabled('TREMISEREF',       lBoExportCFONB) ;
  SetControlEnabled('REMISEDATE',       lBoExportCFONB) ;
  SetControlEnabled('TREMISEDATE',      lBoExportCFONB) ;
  SetControlEnabled('FRAISIMPUTATION',  lBoExportCFONB) ;
  SetControlEnabled('TFRAISIMPUTATION', lBoExportCFONB) ;

  // Spécif esp
  SetControlEnabled('IDCEDENTE',        (lBoExportCFONB) and (TobScenario.GetValue('CPG_CFONBFORMAT')='N32')) ;
  SetControlEnabled('LblIDCEDENTE',     lBoExportCFONB) ;
  SetControlEnabled('EFFETSPHYS',       (lBoExportCFONB) and (TobScenario.GetValue('CPG_CFONBFORMAT')='N32')) ;

end;

procedure TOF_CPMULENCADECA.GestionAccesOngletExport;
begin
  // Onglet edition / export visible ?
  if VH^.PaysLocalisation<>CodeISOES then
    TTabSheet(GetControl('PEXPORT')).TabVisible := False
  else
    begin
    TTabSheet(GetControl('PEXPORT')).TabVisible := ( Operation.Code <> '' ) and ( Operation.AccesCFONB <> taJamais ) ;
    GestionAccesExport ;
    end ;
end;

function TOF_CPMULENCADECA.GetConditionSuiviTreso: String;
begin
  if GetControlText('SUIVITRESO')='X'
    then Result := 'G_SUIVITRESO IN ("' + GetControlText('FLUX')  + '","MIX")'
    else Result := 'G_SUIVITRESO = "' + GetControlText('FLUX') + '"' ;
end;

procedure TOF_CPMULENCADECA.GestionAccesNumBor;
var lBoAvecNumBor : Boolean ;
begin
  lBoAvecNumBor := ( TobScenario <> nil ) and ( TobScenario.GetValue('CPG_BORDEREAUEXP')='X' ) ;
  SetControlVisible('TNUMBORDEREAU', lBoAvecNumBor ) ;
  SetControlVisible('NUMBORDEREAU',  lBoAvecNumBor ) ;
end;

procedure TOF_CPMULENCADECA.MultiDossierChanged(Sender: TObject);
begin
  GestionAccesScenario ;
end;

function TOF_CPMULENCADECA.ConstitueOrigineMS: TOB;
Var lTobEcr    : Tob ;
    i          : Integer ;
    lQEcr      : TQuery ;
    lStDossier : String ;
begin

  // Libération si besoin
  Result        := TOB.Create( 'TOB_ORIGINE', nil, -1 ) ;

{$IFNDEF EAGLCLIENT}
  if FListe.AllSelected then
    begin
    Q.First ;
    While not Q.Eof do
      begin
      if GetControlText('MULTIDOSSIER')<>''
        then lStDossier := GetField('SYSDOSSIER')
        else lStDossier := '' ;

      lQEcr := OpenSelect( ReqEcriture + CleEcriture , lStDossier ) ;
      if not lQEcr.Eof then
        begin
        lTobEcr := TOB.Create('ECRITURE', Result, -1 ) ;
        lTobEcr.SelectDB( '' , lQEcr ) ;
        if (lStDossier <> '') and (lTobEcr.GetNumChamp('SYSDOSSIER') < 1) then
          lTobEcr.AddChampSupValeur( 'SYSDOSSIER', lStDossier ) ;
        end ;
      Ferme( lQEcr ) ;

      Q.Next ;

      end ;
    end
  else
{$ENDIF EAGLCLIENT}
    // Chargement à la ligne
    for i:=0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookmark(i) ;
      {$IFDEF EAGLCLIENT}
      Q.TQ.Seek( FListe.row - 1 ) ;
      {$ENDIF}

      if GetControlText('MULTIDOSSIER')<>''
        then lStDossier := GetField('SYSDOSSIER')
        else lStDossier := '' ;

      lQEcr := OpenSelect( ReqEcriture + CleEcriture , lStDossier ) ;
      if not lQEcr.Eof then
        begin
        lTobEcr := TOB.Create('ECRITURE', Result, -1 ) ;
        lTobEcr.SelectDB( '' , lQEcr ) ;
        if (lStDossier <> '') and (lTobEcr.GetNumChamp('SYSDOSSIER') < 1) then
          lTobEcr.AddChampSupValeur( 'SYSDOSSIER', lStDossier ) ;
        end ;
      Ferme( lQEcr ) ;

      end ;

end;

procedure TOF_CPMULENCADECA.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  if ( csDestroying in Ecran.ComponentState ) then Exit ;

  case Key of

   VK_F10 : begin
            BOuvrirClick(nil);
            Key := 0;
            end;

    end; // case

  if Assigned( FOldKeyDown ) then
     FOldKeyDown(Sender, Key, Shift);

end;

procedure TOF_CPMULENCADECA.GestionAccesEtab;
var lBoAvecEtab : Boolean ;
begin
  lBoAvecEtab := ( TobScenario <> nil ) and ( TobScenario.GetValue('CPG_MULTIETAB')='X' )
                                        and ( TobScenario.GetValue('CPG_SELECTETAB')='FAC' ) ;
  SetControlVisible('TETABLISSEMENTPAYEUR', lBoAvecEtab ) ;
  SetControlVisible('ETABLISSEMENTPAYEUR', lBoAvecEtab ) ;
end;

procedure TOF_CPMULENCADECA.SelectAllClick(Sender: TObject);
var Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULENCADECA.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


function TOF_CPMULENCADECA.DatesOk: Boolean;
Var DD  : TDateTime ;
    Err : integer ;
begin
  Result := FALSE ;

  if Not IsValidDate( FDateComptable.Text ) then
    begin
    HM.Execute(12, Ecran.caption, '') ;
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
      HM.Execute(16+Err, Ecran.caption,'') ;
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
          HM.Execute(12, Ecran.caption,'') ;
          SetActiveTabSheet('PCRITERE') ;
          FDateEche.SetFocus ;
          Exit;
          end
        else
          if not NbJoursOk( DD, StrToDate( FDateEche.Text ) ) then
            begin
            HM.Execute(11, Ecran.caption,'') ;
            SetActiveTabSheet('PCRITERE') ;
            FDateEche.SetFocus ;
            Exit;
            end ;
      end ;
    end ;

  Result := TRUE ;

end;

{JP 04/06/07 : FQ 19747 : Contrôle du filtre pour désactiver certains évènements, notamment
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULENCADECA.ComboFiltreChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FLoadFiltre := True;
  FFiltreChange(Sender);
  FLoadFiltre := False;
end;

{JP 07/06/07 : FQ 17001 : Gestion des banques prévisionnelles
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULENCADECA.InitBanquePrevi;
{---------------------------------------------------------------------------------------}
var
  Pop : TPopupMenu;
begin
  Pop := TPopupMenu(GetControl('POPUPMENU'));
  if Assigned(Pop) then
    if Assigned(Pop.Items[0]) then begin
      Pop.Items[0].Visible := True;
      Pop.Items[0].OnClick := AffecteBanqueClick;
      AddMenuPop(Pop, '', '');
    end;
  SetPlusBanqueCp(GetControl('E_BANQUEPREVI'));
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMULENCADECA.GetChamp(Q : THQuery; sChamp : string) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if Q = nil then Result := GetField(sChamp)
             else Result := TFMul(Ecran).Q.FindField(sChamp).AsVariant;
end;

{JP 07/06/07 : FQ 17001 : Gestion des banques prévisionnelles
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULENCADECA.AffecteBanqueClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  L : TStringList; {FQ TRESO 10294}

    {----------------------------------------------------------------}
    procedure _CreerTob(Q : THQuery = nil);
    {----------------------------------------------------------------}
    var
      F : TOB;
    begin
      {JP 27/09/05 : FQ TRESO 10294 : On ne traite pas les écritures de réglement}
      if L.IndexOf(GetChamp(Q, 'E_JOURNAL')) = -1 then begin
        F := TOB.Create('$$$$', T, -1);
        F.AddChampSupValeur('E_EXERCICE'     , GetChamp(Q, 'E_EXERCICE'));
        F.AddChampSupValeur('E_JOURNAL'      , GetChamp(Q, 'E_JOURNAL'));
        F.AddChampSupValeur('E_DATECOMPTABLE', GetChamp(Q, 'E_DATECOMPTABLE'));
        F.AddChampSupValeur('E_NUMEROPIECE'  , GetChamp(Q, 'E_NUMEROPIECE'));
        F.AddChampSupValeur('E_NUMLIGNE'     , GetChamp(Q, 'E_NUMLIGNE'));
        F.AddChampSupValeur('E_NUMECHE'      , GetChamp(Q, 'E_NUMECHE'));
        F.AddChampSupValeur('E_QUALIFPIECE'  , GetChamp(Q, 'E_QUALIFPIECE'));
        {Pour la mise à jour dans CPAFFECTEBQPREVI_TOF
        JP 21/07/06 : Modification de la fonction WhereEcritureTob : prefixe est devenu inutile
        F.AddChampSupValeur('PREFIXE', 'E');}
      end;
    end;

    {JP 27/09/05 : FQ TRESO 10294 : Chargement des journaux de banque
    {----------------------------------------------------------------}
    function _ChargeListeJournaux : TStringList;
    {----------------------------------------------------------------}
    var
      Q : TQuery;
    begin
      Result := TStringList.Create;
      Q := OpenSQl('SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL = "BQE"', True);
      try
        while not Q.EOF do begin
          Result.Add(Q.FindField('J_JOURNAL').AsString);
          Q.Next;
        end;
      finally
        Ferme(Q);
      end;
    end;

var
  n : Integer;
begin
  {Soit rien n'est sélectionné, soit on est en mode inversé et tut est sélectionné}
  if not TFmul(Ecran).FListe.AllSelected and (TFmul(Ecran).FListe.nbSelected = 0) then begin
    HShowMessage('1;' + Ecran.Caption + ';Veuillez sélectionner au moins une ligne.;W;O;O;O;', '', '');
    Exit;
  end;

  {JP 27/09/05 : FQ TRESO 10294 : Chargement des journaux de banque afin de pouvoir empécher
                 l'affectation de la banque prévisionnelle sur les écritures de réglement}
  L := _ChargeListeJournaux;

  {Constitution de la Tob contenant la clef des écritures sélectionnées}
  T := TOB.Create('$$$$', nil, -1);
  try
    if TFmul(Ecran).FListe.AllSelected  then begin
      TFMul(Ecran).Q.First;
      while not TFMul(Ecran).Q.EOF do begin
        _CreerTob(TFMul(Ecran).Q);
        TFMul(Ecran).Q.Next;
      end;
    end

    else begin
      for n := 0 to TFmul(Ecran).FListe.nbSelected - 1 do begin
        TFmul(Ecran).FListe.GotoLeBookmark(n);
        _CreerTob;
      end;
    end;

    TheTob := T;
    {Appel de la fiche d'affectation de la banque prévisionnelle}
    CP_AffecteBqPrevi(';');
    {Rafraîchissement de la liste}
    TFMul(Ecran).BCherche.Click;
  finally
    if Assigned(T) then FreeAndNil(T);
    if Assigned(L) then FreeAndNil(L);
  end;
end;

function TOF_CPMULENCADECA.GetChampsOpti: string;
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

initialization
  registerclasses ( [ TOF_CPMULENCADECA ] ) ;
end.


