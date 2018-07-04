{***********UNITE*************************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 15/05/2002
Modifié le ... : 23/08/2005 - GCO - FQ 16248
Suite ........ : 23/08/2005 - GCO - FQ 11950 et 16467
Description .. : Source TOF de la FICHE : CPCONSGENE ()
Mots clefs ... : TOF;CPCONSGENE
*****************************************************************}
Unit uTofConsGene ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     Forms,
     UTof,
     UTob,
     Sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     ExtCtrls,        // TImage
     HMsgBox,
     Menus,           // TMenuItem, PopUpMenu
     LookUp,          // LookUpList
     HQry,            // HQuery
     Htb97,           // TToolBarButton97, TToolWindow97
     HRichOle,        // THRichEditOle
     HSysMenu,
     Filtre,          // LoadFiltre
     Windows,         // VK_
     Grids,           // TGridDrawState
     PARAMSOC,        // GetParamSocSecur
     AGLInit,         // TheData
     Ent1,            // VH^, TFichierBase
     uTobDebug,
{$IFDEF MODENT1}
     CPTypeCons,
{$ENDIF MODENT1}
{$IFDEF EAGLCLIENT}
     MainEAgl,        // AGLLanceFiche
{$ELSE}
     Fe_Main,         // AGLLanceFiche
     Db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF COMPTA}
     FPRapproCRE_TOF, //Rapprochement Compta/CRE
{$ENDIF}
     uTofViergeMul;   // Fiche Ancetre Vierge MUL

Type

  TInfoGene = record
    General       : string;
    NatureGene    : string;
    QualifQte1    : string;
    QualifQte2    : string;
    CycleRevision : string;

    Lettrable  : Boolean;
    Pointable  : Boolean;
    Immo       : Boolean;
    Icc        : Boolean;
    IccCapital : Boolean;
    Analytique : Boolean;
    AnaSurAxe1 : Boolean;
    AnaSurAxe2 : Boolean;
    AnaSurAxe3 : Boolean;
    AnaSurAxe4 : Boolean;
    AnaSurAxe5 : Boolean;
    Cre        : Boolean;
    Bqe        : Boolean;
    Collectif  : Boolean;
    SaisieTreso: Boolean;
    VisaRevision : Boolean;
    GCD          : boolean ;
    GCDCOMPTA    : boolean ;
  end;


  TOF_CPCONSGENE = Class (TOF_ViergeMul)

      G_General          : THEdit;
      G_General_         : THEdit;
      G_DateDernMvt      : THEdit;
      G_DateDernMvt_     : THEdit;
      G_Ventilable       : TCheckBox;

      G_NatureGene       : THMultiValComboBox;
      GSens              : THValComboBox;
      GSensReel          : THValComboBox;
      GModeSelection     : THValComboBox;
      GCorresp1          : THMultiValComboBox;
      GSituationCompte   : THValComboBox;

      G_CYCLEREVISION    : THEdit;
      TMPGCYCLEREVISION  : THEdit;

      TabRevision        : TTabSheet;
      TabTablesLibres    : TTabSheet;

      BBlocNote          : TToolBarButton97;
      BCyclePrecedent    : TToolBarButton97;
      BCycleSuivant      : TToolBarButton97;

      PopUpComptable      : TPopUpMenu;
      PopUpSpecifique     : TPopUpMenu;
      PopUpAideRevision   : TPopUpMenu;
      PopUpRevInteg       : TPopUpMenu;
      PopUpParametre      : TPopUpMenu;
      PopUpEdition        : TPopUpMenu;
 
      // Révision
      GTypePonderation   : THValComboBox;
      ZCPonderation      : THValComboBox;
      GValeurPonderation : THNumEdit;
      GBasePonderation   : THEdit;

      GTypeVariation     : THValComboBox;
      ZCVariation        : THValComboBox;
      GValeurVariation   : THNumEdit;
      GNatureVariation   : THValComboBox;

      BEffaceRevision    : TToolBarButton97;

      // Ajout GCO - 18/01/2007
      PRevision          : TPanel;

      IMNoteTravail      : TImage;
      IMTableauVariation : TImage;
      IMCommentaire      : TImage;
      IMVisa             : TImage;
      IMNon              : TImage;
      IMOui              : TImage;
      IM1                : TImage;
      IM2                : TImage;
      IM3                : TImage;
      IM4                : TImage;

      LNoteTravail       : THLabel;
      LTableauVariation  : THLabel;
      LCommentaire       : THLabel;
      LVisa              : THLabel;
      LPonderationCycle  : THLabel;
      TLEtatCycle        : THLabel;
      lETatCycle         : THLabel;

      CBTableauVariation : TCheckBox;
      CBNoteTravail      : TCheckBox;
      CBCptSansCycle     : TCheckBox;
      // Fin Révision

      BImprimer          : TToolBarButton97;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    procedure OnKeyDownEcran   (Sender : TObject; var Key: Word; Shift: TShiftState); override ;
    procedure OnRowEnterFListe (Sender : TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean); override;
    procedure OnClickBBlocNote (Sender : TObject ); override;
    procedure OnPopUpPopF11    (Sender : TObject ); override;

    // GCO MBAMF
    procedure OnPostDrawCellFListe (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState); override;
    procedure AfterShow ; override;

    private
      FToutAccorde          : Boolean; // Pour déterminer les droits de modif du USER
      FBoCriterePonderation : Boolean; // Doit on prendre en compte les critères de la Pondéeration
      FBoCritereVariation   : Boolean; // Doit on prendre en compte les critères de la Variation
      FBoOkGereQte          : Boolean; // Gestion de la quantité sur le dossier
      FBoOkRevision         : Boolean; // Gestion de la révision

      FStSelect          : string;
      FStArgumentTOF     : string;
      FStLiasseDossier   : string;

      FColDebitE         : integer;
      FColCreditE        : integer;
      FColDebitS         : integer;
      FColCreditS        : integer;
      FColVariation      : integer;
      FColPourcentage    : integer;
      FColSoldeE         : integer;
      FColSoldeP         : integer;

      FValeurPonderation : Double;
      FValeurVariation   : Double;

      FStPrefixeE : string;
      FStPrefixeP : string;
      FStCodeExoE : string ;
      FStCodeExoP : string ;

      FExoDate : TExoDate;

      FInfoGene : TInfoGene;

      FTobListeCycle : Tob;
      FTobBDSE       : Tob; // Tob sur la balance de situation en cours
      FTobBDSP       : Tob; // Tob sur le balance de situation précédent

      procedure IndiceColDebCre;
      procedure MiseAJourCaptionEcran;
      procedure ChargeFInfoGene;

      procedure InitBalanceSituation;
      function  ChargeBalanceSituation  ( vStCodeBal : string; vBoSurN : Boolean ) : Boolean;

      procedure OnExitG_General         ( Sender : TObject );
      procedure OnElipsisClickG_General ( Sender : TObject );

      // Initalisation des PopUp
      procedure InitAllPopUp            ( vActivation : Boolean );
      procedure OnPopUpComptable        ( Sender : TObject );
      procedure OnPopUpSpecifique       ( Sender : TObject );
      procedure OnPopUpAideRevision     ( Sender : TObject );
      procedure OnPopUpRevInteg         ( Sender : TObject );
      procedure OnPopUpParametre        ( Sender : TObject );
      procedure OnPopUpEdition          ( Sender : TObject );

      procedure InitPopUpComptable      ( vActivation : Boolean ); //
      procedure InitPopUpSpecifique     ( vActivation : Boolean ); //
      procedure InitPopUpAideRevision   ( vActivation : Boolean ); //
      procedure InitPopUpRevInteg       ( vActivation : Boolean ); //
      procedure InitPopUpParametre      ( vActivation : Boolean ); //
      procedure InitPopUpEdition        ( vActivation : Boolean ); //

      procedure OnClickDetailCompte     ( Sender : TObject );
      procedure OnClickImprimer         ( Sender : TObject );

      procedure GetCellCanvasFListe     ( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

// GP le 15/07/2008 EVOL AGL 7.2
//    procedure OnMouseEnterHyperLinkLabel(Sender: TObject; Enter : Boolean);
      procedure OnMouseLeaveHyperLinkLabel(Sender: TObject);
      procedure OnMouseEnterHyperLinkLabel(Sender: TObject);

      // Outils specifiques
      procedure LettrageMClick          ( Sender : TObject ); // Lettrage manuel
      procedure LettrageAClick          ( Sender : TObject ); // Relevé d'identité bancaire
      procedure DelettreMvtExoRef       ( Sender : TObject );
      procedure DelettreTotal           ( Sender : TObject ); // Delettrage Total du Compte
      procedure OnClickPasseJustifSolde ( Sender : TObject ); // Accéder au détail du compte (&justificatif de soldes)
      procedure OnClickGCD              ( Sender : TObject ); // Créer une créance dans GCD
      procedure OnClickTableauGCD       ( Sender : TObject ); // Tableau justificatif GCD

      procedure TransactionDeLettrageTotal;
      procedure TransactionDelettreMvtExoRef;
      procedure PointageGClick          ( Sender : TObject ); // Pointage du compte
      procedure PointageEnCoursClick    ( Sender : TObject ); // Pointage en cours

      procedure ConsAuxiliaireClick     ( Sender : TObject ); // Appel de la consultation des auxiliaires
      procedure OnClickDetailIcc        ( Sender : TObject ); // Détail ICC du compte
      procedure OnClickCalculIcc        ( Sender : TObject ); // Calcul ICC
      procedure OnClickEmpruntCRE       ( Sender : TObject ); // Emprunt CRE
     {$IFDEF COMPTA}
      procedure OnClickRappECRCRE(Sender: TObject);           // Rapprochement CRE/ECR
     {$ENDIF} 
      procedure ListeImmoClick          ( Sender : TObject );
      // Traitements Comptables
      procedure SaisieEcrClick          ( Sender : TObject ); // Saisie d'une Piece
      procedure SaisieBorClick          ( Sender : TObject ); // Saisie d'un Bordereau
      procedure SaisieTresoClick        ( Sender : TObject ); // Saisie de Treso
      procedure OnClickCutOff           ( Sender : TObject ); // Ecriture de cutoff

      // Fonctions complementaire
      procedure AnalytiquesClick        ( Sender : TObject ); // Mouvements analytiques
      procedure OnClickCommentCpte      ( Sender : TObject );
      procedure HistoCpteClick          ( Sender : TObject ); // Historique du compte
      procedure CumulsGENEClick         ( Sender : TObject ); // Cumul du compte générale
      procedure OnClickRapproComptaImmo ( Sender : TObject ); // Rapprochement Comptabilité / Immo
      procedure OnClickRapproComptaGCD  ( Sender : TObject ); // Rapprochement Comptabilité / GCD

      procedure OnClickCompteGeneral    ( Sender : TObject );
      procedure OnClickCompteRib        ( Sender : TObject );

      // Sous Menu Liasse Fiscale
      procedure OnClickDonneesComp      ( Sender : TObject ); // Données complémentaires
      procedure OnClickInfoLiasse       ( Sender : TObject ); // information de l'affectation du compte dans les liasses
      procedure OnClickTotalCptLiasse   ( Sender : TObject ); // Affectation des montants

      procedure OnClickVisaCompte       ( Sender : TObject ); // Visa du compte (Alt + V)
      procedure OnClickFeuilleExcel     ( Sender : TObject ); // Feuille Excel

      // Sous Menu Suivi de la révision
      procedure OnClickDocPermanent     ( Sender : TObject ); // Document permanent
      procedure OnClickDocAnnuel        ( Sender : TObject ); // Document annuel
      procedure OnClickRechercheDoc     ( Sender : TObject ); // Recherche de document

      // Révision
      procedure OnClickNoteTravail       ( Sender : Tobject ); // Note de travail
      procedure OnClickTableauVariation  ( Sender : Tobject ); // Tableau des variations
      procedure OnClickRechercheEcritures( Sender : Tobject ); // Recherche d'écritures
      procedure OnClickCommentaire       ( Sender : Tobject ); // Commentaire
      procedure OnClickEtatCycle         ( Sender : Tobject ); // Etat du cycle de révision

{$IFDEF COMPTA}
{$IFNDEF CCMP}
      procedure OnClickDRF              ( Sender : TObject ); // Détermination du résultat fiscal
      procedure RepartitionAnaClick     ( Sender : TObject ); // Mouvements analytiques
      procedure OnClickProgrammeTravail ( Sender : TObject ); // Programme de travail
      procedure OnClickAPG              ( Sender : TObject ); // Appréciation générale
      procedure OnClickSCY              ( Sender : TObject ); // Synthèse des cycles
      procedure OnClickEXP              ( Sender : TObject ); // Expert Comptable
      procedure OnClickControleTva      ( Sender : TObject ); // Contrôle de TVA
{$ENDIF}
{$ENDIF}

      // Sous Menu Documentation des travaux
      procedure OnClickMemoCycle        ( Sender : TObject ); // Mémo cycle de révision de la documentation des travaux
      procedure OnClickMemoObjectif     ( Sender : TObject ); // Mémo Objectif de révision de la documentation des travaux
      procedure OnClickMemoSynthese     ( Sender : TObject ); // Mémo synthèse
      procedure OnClickMemoMillesime    ( Sender : TObject ); // Mémo millesime
      procedure OnClickMemoCompte       ( Sender : TObject ); // Mémo compte général

        // Editions
      procedure OnClickJustifSolde      ( Sender : TObject ); // Justifcatif de Solde
      procedure OnClickGLGene           ( Sender : TObject ); // Grand livre général ( OK )
      procedure OnClickGLGeneParQte     ( sender : TObject ); // Grand Livre général par Quantité
      procedure OnClickGLGeneParAuxi    ( Sender : TObject ); // Grand livre général par auxiliaire ( OK )

      procedure OnClickGLAnaSurAxe1     ( Sender : TObject ); // Grand Livre analytique sur Axe 1
      procedure OnClickGLAnaSurAxe2     ( Sender : TObject ); // Grand Livre analytique sur Axe 2
      procedure OnClickGLAnaSurAxe3     ( Sender : TObject ); // Grand Livre analytique sur Axe 3
      procedure OnClickGLAnaSurAxe4     ( Sender : TObject ); // Grand Livre analytique sur Axe 4
      procedure OnClickGLAnaSurAxe5     ( Sender : TObject ); // Grand Livre analytique sur Axe 5
      procedure LancementGLGENEPARANA   ( vFichierBase : TFichierBase );

      procedure OnClickBalGene          ( Sender : TObject ); // Balance générale
      procedure OnClickBalGeneParAuxi   ( Sender : TObject ); // Balance générale par auxiliaire ( OK )
      procedure OnClickBalAnaSurAxe1    ( Sender : TObject ); // Balance generale par section analytique sur Axe 1
      procedure OnClickBalAnaSurAxe2    ( Sender : TObject ); // Balance generale par section analytique sur Axe 2
      procedure OnClickBalAnaSurAxe3    ( Sender : TObject ); // Balance generale par section analytique sur Axe 3
      procedure OnClickBalAnaSurAxe4    ( Sender : TObject ); // Balance generale par section analytique sur Axe 4
      procedure OnClickBalAnaSurAxe5    ( Sender : TObject ); // Balance generale par section analytique sur Axe 5
      procedure LancementBALGENEPARANA  ( vFichierBase : TFichierBase );

      procedure OnClickNoteCtrlCompte   ( Sender : TObject );


      procedure ANALREVISClick          ( Sender : TObject ); // Analyse complementaire

      // Editions de pointage
      procedure OnClickEtatRapprochement( Sender : TObject ); // Etat rapprochement
      procedure OnClickJustifSoldeBQE   ( Sender : TObject ); // Justificatif de solde bancaire

      // Révision avancée
      procedure OnClickBEffaceRevision         ( Sender : TObject );
      procedure OnElipsisClickGBasePonderation ( Sender : TObject );
      procedure OnChangeGTypePonderation       ( Sender : TObject );
      procedure OnCHangeGBasePonderation       ( Sender : TObject );
      procedure OnChangeGNatureVariation       ( Sender : TObject );
//      procedure OnChangeG_CycleRevision        ( Sender : TObject );

      {$IFDEF COMPTA}
      procedure PrepareArgumentGLG ( vBoJustifSolde : Boolean );
      procedure PrepareArgumentBAL ;
      {$ENDIF}

      function  RecupAutreWhere : string;
      function  RecupWhereSituationCompte : string;
      procedure RecupCritereRevision;
      procedure TripotageGSituationCompte;

      function  CompareOperateur( vValeur, vValeurAComparer : Double; vStOperateur : string ) : Boolean;
      procedure AssigneImage (vImage : TImage; vBoAffiche : Boolean; vInID : integer);

   public

      procedure OnClickBCyclePrecedent( Sender : TObject );
      procedure OnClickBCycleSuivant  ( Sender : TObject );
      procedure OnCkickCBCptSansCycle ( Sender : TObject );

   protected
      procedure InitControl                              ; override; // Init des composants de la fiche
      procedure RemplitATobFListe                        ; override;
      procedure RefreshFListe( vBoFetch : Boolean )      ; override; //
//      procedure InitSelectFiltre(T: TOB)                 ; override;
      function  BeforeLoad : Boolean                     ; override;
      function  AjouteATobFListe( vTob : Tob ) : Boolean ; override;
   end;

function CPLanceFiche_CPCONSGENE( vStParam : string = '' ) : string ;

Implementation

uses
{$IFDEF MODENT1}
    CPVersion,
    CPProcMetier,
    CPProcGen,
{$ENDIF MODENT1}
{$IFDEf EAGLCLIENT}
    UtileAGL,      // LanceEtat
{$ELSE}
    EdtREtat,      // LanceEtat
{$ENDIF}

{$IFDEF AMORTISSEMENT}
    Outils,           // ToNone
    AMLISTE_TOF,      // ConsultationImmo
    ImRapCpt,         // EtatRapprochementCompta
{$ENDIF}

{$IFDEF COMPTA}
  {$IFNDEF CCMP}
    uTOFSaisieEteBac,      // CPLanceFiche_SaisieTresorerie
    uTofCPResultatFiscal,  // LanceFiche_DRF()
    CPREPARTITIONANA_TOF,  // CPLanceFiche_RepartitionAnalytique
    CRevInfoDossier_TOM,   // CPLanceFiche_CPREVINFODOSSIER
  {$ENDIF}
    ULibPointage,         {JP 30/07/07 : MsgPointageSurTreso}
    Constantes,           {JP 30/07/07 : CODENEWPOINTAGE}
    CPGestionCreance ,
    SaisBor,              // SaisieFolio
    Saisie,               // TrouveSaisie
    UTofCutOff,
    Lettrage,             // LettrageManuel
    RappAuto,             // Rapprochement automatique
    uTofPointageEcr,      // CPLanceFiche_Pointage
    CPPOINTAGERAP_TOF,    {JP 12/07/07 : FQ 21045 : nouveau pointage : CPLanceFiche_PointageRappro}
    uTOFPointageMul,      // CPLanceFiche_PointageMul
    uTofCPTableauVar,     // CPLanceFiche_CPTableauVar
    uTofCPNoteTravail,    // CPLanceFiche_CPNoteTravail
    CPRevProgTravail_TOF, // CPLanceFiche_CPREVProgTravail
    CPREVDocTravaux_TOF,  // CPLanceFiche_CPRevDOCTRAVAUX

    // *** Editions ***
    uTofCPGLGene,     // CPLanceFiche_CPGLGENE
    uTofCPGLAna,      // CPLanceFiche_CPGLGeneParAna
    CPBalGen_TOF,     // CPLanceFiche_BalanceGeneral
    CPBalGenAuxi_TOF, // CPLanceFiche_BalanceGenAuxi
    CPBalGenAnal_Tof, // CPLanceFiche_BalanceGenAnal
    CPRapproDet_TOF,  // CC_LanceFicheEtatRapproDet
    CPRAPPRODETV7_Tof,    //CC_LanceFicheEtatRapproDetV7
    CPAVERTNEWRAPPRO_TOF, //OkNewPointage
    CPJUSTBQ_TOF,     // CC_LanceFicheJustifPointage
    CPMulAna_TOF,     // MultiCritereAnaZoom
{$ENDIF}

    uLibRevision,         // MiseAJourVisaCompte

    Cummens,              // CumulCpteMensuel
    LettUtil,             // CExecDelettrage
    uTofConsEcr,          // OperationsSurComptes
    uTofRechercheEcr,     // CPLanceFiche_CPRechercheEcr
    uTofConsAuxi,         // CPLanceFiche_CPCONSAUXI
    uTofCPMulMvt,         // MultiCritereMvt
    uLibEcriture,         // CEstAutoriseDelettrage
    uLibWindows,          // AfficheDBCR
    AGLFctSuppr,          // CPSupprimeListeEnreg
    OuvrFermCpte,         // CPModifieComptes
    CPGeneraux_TOM,       // FicheGene
    CPControleLiasse_TOF, // CPLanceFiche_CPControleLiasse
    CPTotalCptLiasse_TOF, // CPLancefiche_CPTotalCptLiasse
    CPControlTva_TOF,     // CPLanceFiche_CPCONTROLTVA
    BanqueCP_Tom,         // FicheBanqueCP
    FichComm,
    SaiSUtil,             // pour le StrSO
    CritEdt,              // TCritEdt
    uTOFHistoCpte ,       // CC_LanceFicheHistoCpte;
    CalcOle,              // Get_CumulPCL
    uImportBob_TOF,       // CPLanceFiche_CPIMPFICHEXCEL
    SaisComm,             // GetO
    variants,             // unassigned
    ComObj;               // CreateOleObject

const cFI_TABLE    = 'CPCONSGENE';
      cOrder       = 'ORDER BY G_GENERAL';
      cStChampsSup = ',G_CUTOFF CUTOFF,G_TOTDEBE TOTDEBE, G_TOTCREE TOTCREE, ' +
                     'G_TOTDEBP TOTDEBP, G_TOTCREP TOTCREP, ' +
                     'G_TOTDEBS TOTDEBS, G_TOTCRES TOTCRES, ' +
                     'G_NATUREGENE NATUREGENE, G_LETTRABLE LETTRABLE, ' +
                     'G_POINTABLE POINTABLE, G_VENTILABLE VENTILABLE, ' +
                     'G_VENTILABLE1 VENTILABLE1, G_VENTILABLE2 VENTILABLE2, ' +
                     'G_VENTILABLE3 VENTILABLE3, G_VENTILABLE4 VENTILABLE4, ' +
                     'G_VENTILABLE5 VENTILABLE5, G_SENS SENS, ' +
                     'G_VISAREVISION VISAREVISION, G_CYCLEREVISION CYCLEREVISION, ' +
                     'G_QUALIFQTE1 QUALIFQTE1, G_QUALIFQTE2 QUALIFQTE2 ';

      cNS = -999999999.99;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPCONSGENE( vStParam : string = '' ) : string ;
begin
{$IFDEF CCSTD}
  // GCO - 26/10/2004 - FQ 14859
  Result := AGLLanceFiche('CP', 'CPCONSGENE', '', '', vStParam);
{$ELSE}
  if CtxPcl in V_Pgi.PgiContexte then
  begin
    if (VH^.CpExoRef.Code = VH^.Encours.Code) or (VH^.CpExoRef.Code = VH^.Suivant.Code) then
    begin
      Result := AGLLanceFiche('CP', 'CPCONSGENE', '', '', vStParam);
    end
    else
    begin
      PgiInfo('Pour pouvoir utiliser cette fonction, l''exercice de référence doit ' +
              'être l''exercice en cours ou le suivant.', 'Consultation des comptes généraux');
    end;
  end
  else
    Result := AGLLanceFiche('CP', 'CPCONSGENE', '', '', vStParam);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/02/2007
Modifié le ... : 17/04/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClose ;
begin
  FreeAndNil( FTobListeCycle );
  FreeAndNil( FTobBDSE );
  FreeAndNil( FTobBDSP );
  Inherited ;
end ;

procedure TOF_CPCONSGENE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPCONSGENE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPCONSGENE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPCONSGENE.IndiceColDebCre ;
var
 lStChamps : string;
 lStListe  : string;
 lInIndex  : integer;
begin
  lStListe        := FStListeChamps;
  lInIndex        := 1;
  FColDebitE      := -1;
  FColCreditE     := -1;
  FColDebitS      := -1;
  FColCreditS     := -1;
  FColSoldeP      := -1;
  FColSoldeE      := -1;
  FColVariation   := -1;
  FColPourcentage := -1;

  while lStListe <> '' do
  begin
   lStChamps := READTOKENST(lStListe);
   if ( lStChamps = 'G_TOTDEBE' ) then
     FColDebitE := lInIndex
   else
     if ( lStChamps = 'G_TOTCREE' ) then
       FColCreditE := lInIndex
     else
       if ( lStChamps = 'G_TOTDEBS' ) then
         FColDebitS := lInIndex
       else
         if ( lStChamps = 'G_TOTCRES' ) then
           FColCreditS := lInIndex
         else
           if ( lStChamps = 'SOLDEP' ) then
             FColSoldeP := lInIndex
           else
             if ( lStChamps = 'SOLDEE' ) then
               FColSoldeE := lInIndex
             else
               if ( lStChamps = 'VARIATION' ) then
                 FColVariation := lInIndex
               else
                 if ( lStChamps = 'POURCENTAGE' ) then
                 FColPourcentage := lInIndex;
               
    Inc(lInIndex);
  end; // while
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/05/2002
Modifié le ... : 22/01/2003
Description .. : - LG - 22/01/2003 - appel de init auto pour gere
Suite ........ : l'autocomplementation
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnArgument (S : String ) ;
begin
  FFI_Table              := cFI_TABLE;
  FStListeParam          := 'CPCONSGENE';
  FStArgumentTOF         := S;

  inherited ;
  Ecran.HelpContext := 7603100;

  // Elements de la fiche Ancêtre
  FListe.OnDblClick      := OnClickDetailCompte;
  FListe.PostDrawCell    := OnPostDrawCellFListe;
  //  FListe.OnFlipSelection := OnFlipSelectionFListe;
  FListe.Options         := FListe.Options + [goRangeSelect] ;
  FListe.MultiSelect     := False; // GCO - 28/07/2005

  IndiceColDebCre;
  FStSelect    := FStListeChamps;
  FStSelect    := CSqlTextFromList(FStSelect) ;

  Ecran.OnKeyDown := OnKeyDownEcran;

  G_General       := THEdit(GetControl('G_GENERAL', True));
  G_General_      := THEdit(GetControl('G_GENERAL_', True));
  G_DateDernMvt   := THEdit(GetControl('G_DATEDERNMVT', True));
  G_DateDernMvt_  := THEdit(GetControl('G_DATEDERNMVT_', True));
  G_Ventilable    := TCheckBox(GetControl('G_VENTILABLE', True));

  GSens           := THValComboBox(GetControl('G_SENS', True));
  GSensReel       := THValComboBox(GetControl('GSENSREEL', True));
  GModeSelection  := THValComboBox(GetControl('GMODESELECTION', True));
  GCorresp1       := THMultiValComboBox(GetControl('GCORRESP1', True));
  GSituationCompte:= THValComboBox(GetControl('GSITUATIONCOMPTE', True));

  G_NatureGene    := THMultiValComboBox(GetControl('G_NATUREGENE', True));

  G_CycleRevision := THEdit(GetControl('G_CYCLEREVISION', True));
  TMPGCYCLEREVISION := THEdit(GetControl('TMPGCYCLEREVISION', True));

  BBlocNote       := TToolBarButton97(GetControl('BBLOCNOTE', True));
  BCyclePrecedent := TToolBarButton97(GetControl('BCYCLEPRECEDENT', True));
  BCycleSuivant   := TToolBarButton97(GetControl('BCYCLESUIVANT', True));

  PopUpComptable     := TPopUpMenu(GetControl('PopUpComptable', True));
  PopUpSpecifique    := TPopUpMenu(GetControl('PopUpSpecifique', True));
  PopUpAideRevision  := TPopUpMenu(GetControl('PopUpAideRevision', True));
  PopUpRevInteg      := TPopUpMenu(GetControl('PopUpRevInteg', True));
  PopUpParametre     := TPopUpMenu(GetControl('PopUpParametre', True));
  PopUpEdition       := TPopUpMenu(GetControl('PopUpEdition', True));

  TabRevision        := TTabSheet(GetControl('TABREVISION', True));
  TabTablesLibres    := TTabSheet(GetControl('TABTABLESLIBRES', True));

  // Révision avancée
  GTypePonderation   := THValComboBox(GetControl('GTYPEPONDERATION', True));
  ZCPonderation      := THValComboBox(GetControl('ZCPONDERATION', True));
  GValeurPonderation := THNumEdit(GetControl('GVALEURPONDERATION', True));
  GBasePonderation   := THEdit(GetControl('GBASEPONDERATION', True));

  GTypeVariation     := THValComboBox(GetControl('GTYPEVARIATION', True));
  ZCVariation        := THValComboBox(GetControl('ZCVARIATION', True));
  GValeurVariation   := THNumEdit(GetControl('GVALEURVARIATION', True));
  GNatureVariation   := THValComboBox(GetControl('GNATUREVARIATION', True));

  BEffaceRevision    := TToolBarButton97(GetControl('BEFFACEREVISION', True));
  BCyclePrecedent    := TToolBarButton97(GetControl('BCYCLEPRECEDENT', True));
  BCycleSuivant      := TToolBarButton97(GetControl('BCYCLESUIVANT', True));

  PRevision          := TPanel(GetControl('PREVISION', True));

  IMNoteTravail      := TImage(GetControl('IMNOTETRAVAIL', True));
  IMTableauVariation := TImage(GetControl('IMTABLEAUVARIATION', True));
  IMCommentaire      := TImage(GetControl('IMCOMMENTAIRE', True));
  IMVisa             := TImage(GetControl('IMVISA', True));
  IMNOn              := TImage(GetControl('IMNON', True));
  IMOUI              := TImage(GetControl('IMOUI', True));
  IM1                := TImage(GetControl('IM1', True));
  IM2                := TImage(GetControl('IM2', True));
  IM3                := TImage(GetControl('IM3', True));
  IM4                := TImage(GetControl('IM4', True));

  LNoteTravail       := THLabel(GetControl('LNOTETRAVAIL', True));
  LTableauVariation  := THLabel(GetControl('LTABLEAUVARIATION', True));
  LCommentaire       := THLabel(GetControl('LCOMMENTAIRE', True));
  LVisa              := THLabel(GetControl('LVISA', True));
  LPonderationCycle  := THLabel(GetControl('LPONDERATIONCYCLE', True));
  TLEtatCycle        := THLabel(GetControl('TLETATCYCLE', True));
  lEtatCycle         := THLabel(GetControl('LETATCYCLE', True));

  CBTableauVariation := TCheckBox(GetControl('CBTABLEAUVARIATION', True));
  CBNoteTravail      := TCheckBox(GetControl('CBNOTETRAVAIL', True));
  CBCptSansCycle     := TCheckBox(GetControl('CBCPTSANSCYCLE', True));
  // Fin de Révision avancée

  BImprimer          := TToolBarButton97(GetControl('BIMPRIMER', True));

//  FListeByUser.OnSelect := InitSelectFiltre;

  //----------------------- Branchement des événements -------------------------
  G_Ventilable.Enabled    := not (((ctxPCL in V_PGI.PGIContexte) {or (EstComptaSansAna)}) and (GetParamSocSecur('SO_CPPCLSANSANA', True)=TRUE));

  PopF11.OnPopup            := OnPopUpPopF11;
  PopUpComptable.OnPopup    := OnPopUpComptable;
  PopUpSpecifique.OnPopup   := OnPopUpSpecifique;
  PopUpAideRevision.OnPopup := OnPopUpAideRevision;
  PopUpRevInteg.OnPopup     := OnPopUpRevInteg;
  PopUpParametre.OnPopup    := OnPopUpParametre;
  PopUpEdition.OnPopUp      := OnPopUpEdition;

  FListe.GetCellCanvas      := GetCellCanvasFListe;
  FListe.OnRowEnter         := OnRowEnterFListe;

  GTypePonderation.OnChange := OnChangeGTypePonderation;
  GBasePonderation.OnChange := OnChangeGBasePonderation;
  GBasePonderation.OnElipsisClick := OnElipsisClickGBasePonderation;
  GNatureVariation.OnChange := OnChangeGNatureVariation;
//  G_CycleRevision.OnChange  := OnChangeG_CycleRevision;

  BEffaceRevision.OnClick := OnClickBEffaceRevision;
  BCyclePrecedent.OnClick := OnClickBCyclePrecedent;
  BCycleSuivant.OnClick   := OnClickBCycleSuivant;
  CBCptSansCycle.OnClick  := OnCkickCBCptSansCycle;

  TripotageGSituationCompte;

  G_General.MaxLength       := VH^.Cpta[fbGene].lg;
  G_General_.MaxLength      := VH^.Cpta[fbGene].lg;
  G_General.OnElipsisClick  := OnElipsisClickG_General;
  G_General_.OnElipsisClick := OnElipsisClickG_General;
  G_General.OnExit          := OnExitG_General;
  G_General_.OnExit         := OnExitG_General;
  BImprimer.OnClick         := OnClickImprimer;


// GP le 15/07/2008 EVOL AGL 7.2
(*
  LNoteTravail.OnMouseEnter      := OnMouseEnterHyperLinkLabel;
  LTableauVariation.OnMouseEnter := OnMouseEnterHyperLinkLabel;
  LCommentaire.OnMouseEnter      := OnMouseEnterHyperLinkLabel;
  TLEtatCycle.OnMouseEnter       := OnMouseEnterHyperLinkLabel;
*)
  LNoteTravail.OnMouseEnter      := OnMouseEnterHyperLinkLabel;
  LTableauVariation.OnMouseEnter := OnMouseEnterHyperLinkLabel;
  LCommentaire.OnMouseEnter      := OnMouseEnterHyperLinkLabel;
  TLEtatCycle.OnMouseEnter       := OnMouseEnterHyperLinkLabel;
  LNoteTravail.OnMouseLeave      := OnMouseLeaveHyperLinkLabel;
  LTableauVariation.OnMouseLeave := OnMouseLeaveHyperLinkLabel;
  LCommentaire.OnMouseLeave      := OnMouseLeaveHyperLinkLabel;
  TLEtatCycle.OnMouseLeave       := OnMouseLeaveHyperLinkLabel;


  LNoteTravail.OnClick           := OnClickNoteTravail;
  LTableauVariation.OnClick      := OnClickTableauVariation;
  LCommentaire.OnClick           := OnClickCommentaire;
  TLEtatCycle.OnClick            := OnClickEtatCycle;
  lEtatcycle.OnClick             := OnClickEtatCycle;

  // GCO - 16/11/2005
  FBoOkGereQte := GetParamSocSecur('SO_CPPCLSAISIEQTE', False);

  if CtxPcl in V_Pgi.PgiContexte then
  begin
    FExoDate.Code := VH^.CPExoRef.Code;
    FExoDate.Deb  := VH^.CPExoRef.Deb;
    FExoDate.Fin  := VH^.CPExoRef.Fin;
  end
  else
  begin
    if (VH^.Entree.Code = VH^.EnCours.Code) or
       (VH^.Entree.Code = VH^.Suivant.Code) then
    begin
      FExoDate.Code := VH^.Entree.Code;
      FExoDate.Deb  := VH^.Entree.Deb;
      FExoDate.Fin  := VH^.Entree.Fin;
    end
    else
    begin
      FExoDate.Code := VH^.EnCours.Code;
      FExoDate.Deb  := VH^.EnCours.Deb;
      FExoDate.Fin  := VH^.EnCours.Fin;
    end;
  end;

  FStLiasseDossier := GetParamSocSecur('SO_CPCONTROLELIASSE', '');

  // Paramétres de la révision
  FBoOkRevision  := VH^.Revision.Plan <> '';
  FTobListeCycle := Tob.Create('LISTECYCLE', nil, -1);

  if GetParamSocSecur('SO_CPGSI', False, True) then
  begin
    InitBalanceSituation;
  end;
  
  InitControl; // Valeur par défaut de tous les controles de la fiche

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/06/2002
Modifié le ... :   /  /
Description .. : Initialise les composants dans leur état par défaut à l' ouverture
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.InitControl;
begin
  // Init des valeurs
  G_General.Text         := '';
  G_General_.Text        := '';

  G_NatureGene.Value     := '<<' + TraduireMemoire('Tous') + '>>';
  G_CycleRevision.Plus   := ' AND CCY_EXERCICE = "' + VH^.EnCours.Code + '"';

  GModeSelection.Value   := 'ALL';
  GSensReel.Value        := '0';
  GSens.Value            := '';
  GsituationCompte.Value := '';

  G_DateDernMvt.Text     := DateToStr( iDate1900 );
  G_DateDernMvt_.Text    := DateToStr( iDate2099 );

  GCorresp1.Value        := '<<' + TraduireMemoire('Tous') + '>>';

  // Gestion de la révision
  // GCO - 11/06/2007 - FQ 20584
  PRevision.Visible := FBoOkRevision and VH^.OkModRIC and JaiLeRoleCompta( RcReviseur );
  if not FBoOkRevision then
  begin
    G_CycleRevision.Enabled    := False;
    BCyclePrecedent.Enabled    := False;
    BCycleSuivant.Enabled      := False;
    CBTableauVariation.Enabled := False;
    CBTableauVariation.State   := CbGrayed;
    CBNoteTravail.Enabled      := False;
    CBNoteTravail.State        := CbGrayed;
    CBCptSansCycle.Enabled     := False;
    CBCptSansCycle.State       := CbGrayed;
  end;

  OnChangeGTypePonderation(nil);
  OnChangeGBasePonderation(nil);
  OnClickBEffaceRevision(nil);

  // GCO - 25/10/2006 - FQ 16894, FQ 18087
  if ((FExoDate.Code = VH^.Suivant.Code) and (VH^.EnCours.Code = '')) or
     ((FExoDate.Code = VH^.Encours.Code) and (VH^.Precedent.Code = '')) then
  begin
    GTypeVariation.Enabled   := False;
    ZCVariation.Enabled      := False;
    GValeurVariation.Enabled := False;
    GNatureVariation.Enabled := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnLoad ;
var lStArgument : string;
    lStTemp     : string;
begin
  LibellesTableLibre(TabTablesLibres, 'TG_TABLE', 'G_TABLE', 'G');
  
  Inherited ;

  if FStArgumentTOF <> '' then
  begin
    // Compte général
    lStArgument := FStArgumentTOF;
    G_General.Text := ReadTokenSt(lStArgument);

    // Cycle de révision
    lStTemp := ReadTokenSt(lStArgument);
    if lStTemp <> '' then
    begin
      lStTemp := FindEtReplace( lStTemp, '||', ';', True );
      G_CycleRevision.Text := lStTemp;
    end;
  end;

  InitAllPopUp( False );
  FToutAccorde := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False);
  InitAutoSearch ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.AfterShow;
var lNbMoisSoldeE : word   ;
    lNbMoisSoldeP : word   ;
    lPremMois     : word   ;
    lPremAnnee    : word   ;
begin
  inherited;
  if G_General.CanFocus then
    G_General.SetFocus;

  // GCO - 26/09/2007 - FQ 21519
  if FTobBDSE <> nil then
  begin
    NombreMois( FTobBDSE.GetDateTime('BSI_DATE1'),
                FTobBDSE.GetDateTime('BSI_DATE2'), lPremMois, lPremAnnee, lNbMoisSoldeE)
  end
  else
  begin
    if FExoDate.Code = VH^.Suivant.Code then
      NombreMois( VH^.Suivant.Deb, VH^.Suivant.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE)
    else
      NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
  end;

  if FTobBDSP <> nil then
    NombreMois( FTobBDSP.GetDateTime('BSI_DATE1'),
                FTobBDSP.GetDateTime('BSI_DATE2'), lPremMois, lPremAnnee, lNbMoisSoldeP)
  else
  begin
    if FExoDate.Code = VH^.Suivant.Code then
      NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP)
    else
      NombreMois( VH^.Precedent.Deb, VH^.Precedent.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);
  end;

  if (lNbMoisSoldeE = lNbMoisSoldeP) then
  begin
    GTypeVariation.Enabled := False;
    GTypevariation.Value   := 'ABSOLUE';
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/05/2002
Modifié le ... :   /  /
Description .. : Charge la Tob avec la requete grille avec la requête
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.RemplitATobFListe;
var lStSelect      : string ;
    lStSelectRev   : string ; // Select des champs de la révision
    lStLeftJoinRev : string ; // Ajout des LEFT JOIN de la révision
    lStWhere       : string ;
    lSt            : string ;
begin
  if (FExoDate.Code = VH^.Encours.Code) or (FExoDate.Code = VH^.Suivant.Code) then
  begin
    lSt := RecupWhereCritere(PageControl);
    if lSt <> '' then
      lStWhere := lSt + RecupWhereSituationCompte + RecupAutreWhere
    else
      lStWhere := ' WHERE ' + RecupWhereSituationCompte + RecupAutreWhere;

    lStSelectRev   := '';
    lStLeftJoinRev := '';

    if FBoOkRevision then
    begin
      lStSelectRev := ', ' +
                     'IIF((CTV_LIGNE IS NULL), "-", "X") TABLEAUVARIATION, ' +
                     'IIF((CNO_LIGNE IS NULL), "-", "X") NOTETRAVAIL, ' +
                     'IIF((CBN_CODE IS NULL), "-", "X") MILLESIME ';

      lStLeftJoinRev := 'LEFT JOIN CPTABLEAUVAR ON CTV_GENERAL = G_GENERAL AND ' +
                        'CTV_EXERCICE = "' + VH^.EnCours.Code + '" AND CTV_LIGNE = 1 ' +
                        'LEFT JOIN CPNOTETRAVAIL ON CNO_GENERAL = G_GENERAL AND ' +
                        'CNO_EXERCICE = "'+ VH^.EnCours.Code + '" AND CNO_LIGNE = 1 ' +
                        'LEFT JOIN CREVCYCLE ON CCY_CODECYCLE = G_CYCLEREVISION AND ' +
                        // GCO - 09/05/2007 - FQ 20288
                        'CCY_EXERCICE = "' + VH^.EnCours.Code + '" ' +
                        'LEFT JOIN CREVBLOCNOTE ON CBN_CODE = G_GENERAL AND ' +
                        'CBN_NATURE = "GEN" AND CBN_EXERCICE = "' + VH^.EnCours.Code + '" ';
    end;                  

    // Cas particulier pour la requête générée, on recherche les comptes
    // mouvementés sur les deux exercices sans les ANO
    if GModeSelection.Value = 'EXS' then
    begin
      lStSelect := 'SELECT DISTINCT(E_GENERAL),' + CSqlTextFromList(FStListeChamps) + ' ' +
                   cStChampsSup + lStSelectRev + ' FROM ECRITURE ' +
                   'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
                   'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' + lStLeftJoinRev +
                   lStWhere + ' AND E_QUALIFPIECE = "N"' +
                   ' AND (E_EXERCICE = "' + FStCodeExoE + '" OR E_EXERCICE = "' + FStCodeExoP + '")' +
                   ' AND J_NATUREJAL <> "ANO" '  + cOrder;
    end
    else
    begin
      lStSelect := 'SELECT ' + CSqlTextFromList(FStListeChamps) + ' ' +
                   cStChampsSup + lStSelectRev;

      lStSelect := lStSelect + ' FROM GENERAUX ' + lStLeftJoinRev;

      lStSelect := lStSelect + lStWhere + ' ' + cOrder;
    end;

    // On envoie le texte SQL de la requête à la fiche ANCETRE pour qu'elle fasse le OPENSQL
    AStSqlTobFListe := lStSelect;

    // Remplit FRevision avec les valeurs saisies dans l'onglet "Révision Avancée"
    RecupCritereRevision;

  end
  else
  begin
    PgiInfo('Pour pouvoir utiliser cette fonction, l''exercice de référence doit ' +
             'être l''exercice en cours ou le suivant.', 'Consultation des comptes généraux');

    AStSqlTobFListe := '';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/11/2002
Modifié le ... : 20/01/2003
Description .. : Récupère la condition WHERE des critères qui
Suite ........ : n'appartiennent
Suite ........ : pas à la table
Suite ........ : - FB 11804 - la gestion des plans de correspond ne fct pas
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSGENE.RecupAutreWhere : string;
var lStValCorresp1 : string;
    lStLibCorresp1 : string;
begin
  Result := '';

  // Mode de sélection des comptes
  // le Cas où Value = 'EXS' est géré à la création du texte SQL
  if GModeSelection.Value = 'EXT' then
  begin // Comptes mouvementés sur les deux exercices (avec ANO)
    Result := Result + ' AND ((G_TOTDEB' + FStPrefixeE + ' <> 0 OR G_TOTCRE' + FStPrefixeE + ' <> 0) OR' +
                       ' (G_TOTDEB' + FStPrefixeP + ' <> 0 OR G_TOTCRE' + FStPrefixeP + ' <> 0))';
  end
  else
    if GModeSelection.Value = 'NSL' then
    begin // Comptes non soldés sur les deux exercices
      Result := Result + ' AND ((G_TOTDEB' + FStPrefixeE + ' <> G_TOTCRE' + FStPrefixeE + ') OR' +
                         ' (G_TOTDEB' + FStPrefixeP + ' <> G_TOTCRE' + FStPrefixeP + '))';
    end;

  // Sens Réel du solde
  if GSensReel.Value = '1' then
    Result := Result + ' AND (G_TOTDEB' + FStPrefixeE + ' > G_TOTCRE' + FStPrefixeE + ')'
  else
    if GSensReel.Value = '2' then
      Result := Result + ' AND (G_TOTDEB' + FStPrefixeE + ' < G_TOTCRE' + FStPrefixeE + ')'
    else
      if GSensReel.Value = '3' then
        Result := Result + ' AND (G_TOTDEB' + FStPrefixeE + ' = G_TOTCRE' + FStPrefixeE + ')';

  // GCO - 15/09/2004 - FQ 14087 - Gestion du G_CONFIDENTIEL
  if (V_Pgi.Confidentiel = '0') then
  begin
    Result := Result + ' AND ((G_CONFIDENTIEL = "0") OR (G_CONFIDENTIEL = "-")) ';
  end
  else
  begin
    Result := Result + ' AND (((G_CONFIDENTIEL = "-") OR (G_CONFIDENTIEL = "X")) OR ' +
                       ' (G_CONFIDENTIEL <= "' + V_PGI.Confidentiel + '"))';
  end;

  // GCO - 23/08/2005 - FQ 11950 et 16467
  // Compte de Correspondance
  TraductionTHMultiValComboBox( GCorresp1, lStValCorresp1, lStLibCorresp1, 'G_CORRESP1', True );
  if lStValCorresp1 <> '' then
    Result := Result + ' AND ' + lStValCorresp1;

  if not fBoOkRevision then Exit;

  // GCO - 04/04/2007 - Gestion des comptes sans cycle de révision
  if CBCptSansCycle.State <> cbGrayed then
    Result := Result + ' AND G_CYCLEREVISION ' + IIF(CBCptSansCycle.Checked, '= ""', '<> ""');

  // GCO - 15/03/2007 - Gestion des case à cocher Tableau de Variation et Note de Travail
  if CBTableauVariation.State <> CbGrayed then
    Result := Result + ' AND CTV_LIGNE IS ' + IIF( CBTableauVariation.Checked, 'NOT NULL', 'NULL');

  if CBNoteTravail.State <> CbGrayed then
    Result := Result + ' AND CNO_LIGNE IS ' + IIF( CBNoteTravail.Checked, 'NOT NULL', 'NULL');

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSGENE.RecupWhereSituationCompte: string;
begin

  if GSituationCompte.Value = 'SI1' then
  begin // Comptes mouvementés sur l'exercice mais sans mouvements sur le précédent
    Result := ' AND G_GENERAL IN (SELECT DISTINCT (E1.E_GENERAL) FROM ECRITURE E1 WHERE' +
              ' E1.E_EXERCICE = "' + FStCodeExoE + '" AND' +
              ' E1.E_ECRANOUVEAU = "N" AND' +
              ' E1.E_QUALIFPIECE = "N" AND E1.E_GENERAL NOT IN' +
              ' (SELECT E2.E_GENERAL FROM ECRITURE E2 WHERE' +
              ' E2.E_EXERCICE = "' + FStCodeExoP + '" AND' +
              ' E2.E_ECRANOUVEAU = "N" AND' +
              ' E2.E_QUALIFPIECE = "N" ))';
  end
  else
  if GSituationCompte.Value = 'SI2' then
  begin // Comptes sans mouvements sur l'exercice mais mouvementés sur le précédent
    Result := ' AND G_GENERAL IN (SELECT DISTINCT (E1.E_GENERAL) FROM ECRITURE E1 WHERE' +
              ' E1.E_EXERCICE = "' + FStCodeExoP + '" AND' +
              ' E1.E_ECRANOUVEAU = "N" AND' +
              ' E1.E_QUALIFPIECE = "N" AND E1.E_GENERAL NOT IN' +
              ' (SELECT E2.E_GENERAL FROM ECRITURE E2 WHERE' +
              ' E2.E_EXERCICE = "' + FStCodeExoE + '" AND' +
              ' E2.E_ECRANOUVEAU = "N" AND' +
              ' E2.E_QUALIFPIECE = "N" ))';
  end
  else
  if GSituationCompte.Value = 'SI3' then
  begin // Comptes non soldés sur l'exercice mais soldés sur le précédent
    Result := ' AND (G_TOTDEB' + FStPrefixeE + ' <> G_TOTCRE' + FStPrefixeE + ')' +
              ' AND (G_TOTDEB' + FStPrefixeP + ' = G_TOTCRE' + FStPrefixeP + ')';
  end
  else
  if GSituationCompte.Value = 'SI4' then
  begin // Comptes soldés sur l'exercice mais son soldés sur le précédent
    Result := ' AND (G_TOTDEB' + FStPrefixeE + ' = G_TOTCRE' + FStPrefixeE + ')' +
              ' AND (G_TOTDEB' + FStPrefixeP + ' <> G_TOTCRE' + FStPrefixeP + ')';
  end
  else
  if GSituationCompte.Value = 'SI5' then
  begin //Comptes n'ayant pas le même sens sur l'exercice que sur le précédent
    Result := ' AND (((G_TOTDEB' + FStPrefixeE + ' - G_TOTCRE' + FStPrefixeE + ' > 0) AND ' +
                    '(G_TOTDEB' + FStPrefixeP + ' - G_TOTCRE' + FstPrefixeP + ' < 0)) OR ' +
                    '((G_TOTDEB' + FStPrefixeE + ' - G_TOTCRE' + FStPrefixeE + ' < 0) AND ' +
                    '(G_TOTDEB' + FStPrefixeP + ' - G_TOTCRE' + FstPrefixeP + ' > 0)))';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.RecupCritereRevision;
var lTotalRubrique : Double;
    lDate1 : TDateTime;
    lDate2 : TDateTime;
    lBalance  : TStringList;
    lTabloExt : TabloExt;
begin
  // Récupération du montant de la rubrique comptable sélectionée
  if GBasePonderation.Text <> '' then
  begin
    lBalance := TStringList.Create;
    lDate1 := FExoDate.Deb;
    lDate2 := FExoDate.Fin;
    lTotalRubrique := GetCumul('RUB', GBasePonderation.Text, '', '', '', '',
                               FExoDate.Code, lDate1, lDate2, True, True,
                               lBalance, lTabloExt, False);

    // GCO - 07/06/2007 - FQ 20613
    //FValeurPonderation := (lTotalRubrique * Valeur(GValeurPonderation.Text)) / 100;
    FValeurPonderation := (lTotalRubrique * GValeurPonderation.Value) / 100;
    lBalance.Free;
  end
  else
    // GCO - 07/06/2007 - FQ 20613
    //FValeurPonderation := Valeur(GValeurPonderation.Text);
    FValeurPonderation := GValeurPonderation.Value;

  if GNatureVariation.ItemIndex <> -1 then
  begin
    // GCO - 07/06/2007 - FQ 20613
    //FValeurVariation := Valeur(GValeurVariation.Text);
    FValeurVariation := GValeurVariation.Value;
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/04/2007
Modifié le ... : 05/06/2007
Description .. : 
Suite ........ : 
Suite ........ : - LG - 05/06/2007 - FB 20548 -on ouvre les TIC ou TID que 
Suite ........ : pour els racines de compte 411
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.ChargeFInfoGene;
begin
  if ATobRow = nil then
  begin
    FInfoGene.General       := '';
    FInfoGene.NatureGene    := '';
    FInfoGene.QualifQte1    := '';
    FInfoGene.QualifQte2    := '';
    FInfoGene.CycleRevision := '';

    FInfoGene.Lettrable  := False;
    FInfoGene.Pointable  := False;
    FInfoGene.Immo       := False;
    FInfoGene.Icc        := False;
    FInfoGene.IccCapital := False;
    FInfoGene.Analytique := False;
    FInfoGene.AnaSurAxe1 := False;
    FInfoGene.AnaSurAxe2 := False;
    FInfoGene.AnaSurAxe3 := False;
    FInfoGene.AnaSurAxe4 := False;
    FInfoGene.AnaSurAxe5 := False;
    FInfoGene.Cre        := False;
    FInfoGene.Bqe        := False;
    FInfoGene.Collectif  := False;
    FInfoGene.SaisieTreso  := False;
    FInfoGene.VisaRevision := False;
    Exit;
  end;

  // Ici, Recharge systématique des infos qui peuvent être modifiées
  // GCO - 03/07/2007 - FQ 20911
  FInfoGene.VisaRevision := ATobRow.GetString('G_VISAREVISION') = 'X';

  if ATobRow.GetString('G_GENERAL') <> FInfoGene.General then
  begin
    with FInfoGene do
    begin
      General := ATobRow.GetString('G_GENERAL');

      NatureGene := ATobRow.GetString('NATUREGENE');

      QualifQte1 := ATobRow.GetString('QUALIFQTE1');

      QualifQte2 := ATobRow.GetString('QUALIFQTE2');

      Lettrable := ((ATobRow.GetString('NATUREGENE') = 'TIC') or
                    (ATobRow.GetString('NATUREGENE') = 'TID') or
                    (ATobRow.GetString('NATUREGENE') = 'DIV')) and
                    (ATobRow.GetString('LETTRABLE')  = 'X');

      if VH^.PointageJal then
        Pointable := False
      else
        Pointable := (ATobRow.GetString('POINTABLE') = 'X');

      Immo      := (ATobRow.GetString('NATUREGENE') = 'IMO') and (VH^.OkModImmo);
      Collectif := (ATobRow.GetString('NATUREGENE') = 'COF') or (ATobRow.GetString('NATUREGENE') = 'COS') or
                   (ATobRow.GetString('NATUREGENE') = 'COC') or (ATobRow.GetString('NATUREGENE') = 'COD') ;

      if CtxPcl in V_Pgi.PGIContexte then
      begin
        Icc := (VH^.OkModIcc) and Presence('ICCGENERAUX', 'ICG_GENERAL', General) and
               (General <> GetParamSocSecur('SO_ICCCOMPTECAPITAL', ''));

        // GCO - 02/10/2007 - FQ 21509
        IccCapital := (VH^.OkModIcc) and (General = GetParamSocSecur('SO_ICCCOMPTECAPITAL', ''));

        Cre := (VH^.OkModCre) and ((Copy(General, 1, 2) = '16') Or (Copy(General, 1, 4) = '6611'));
      end
      else
      begin
        Icc := Presence('ICCGENERAUX', 'ICG_GENERAL', General); // Sans Seria pour PME
        Cre := (Copy(General, 1, 2) = '16') Or (Copy(General, 1, 4) = '6611');
      end;

      Analytique := (ATobRow.GetString('VENTILABLE')  = 'X');
      AnaSurAxe1 := (ATobRow.GetString('VENTILABLE1') = 'X');
      AnaSurAxe2 := (ATobRow.GetString('VENTILABLE2') = 'X');
      AnaSurAxe3 := (ATobRow.GetString('VENTILABLE3') = 'X');
      AnaSurAxe4 := (ATobRow.GetString('VENTILABLE4') = 'X');
      AnaSurAxe5 := (ATobRow.GetString('VENTILABLE5') = 'X');

      Bqe := (ATobRow.GetString('NATUREGENE') = 'BQE') or
             (ATobRow.GetString('NATUREGENE') = 'TIC') or
             (ATobRow.GetString('NATUREGENE') = 'TID');

      SaisieTreso := ExisteSQL('SELECT J_CONTREPARTIE FROM JOURNAL, BANQUECP WHERE ' +
                               'J_NATUREJAL = "BQE" AND BQ_GENERAL = J_CONTREPARTIE AND ' +
                               'BQ_NODOSSIER="'+ V_PGI.NoDossier + '" AND ' + // 24/10/2006 YMO Multisociétés
                               'BQ_DEVISE = "' + V_PGI.DevisePivot + '" AND ' +
                               'J_CONTREPARTIE = "' + General + '"');

      CycleRevision := ATobRow.GetString('CYCLEREVISION');

      GCD           := VH^.OkModGCD and ( Copy(ATobRow.GetString('G_GENERAL'),1,2) = '41' ) ;
      GCDCOMPTA     := VH^.OkModGCD and
                     ( ( ATobRow.GetString('G_GENERAL') = GetParamSocSecur('SO_CPGCDGENERAL','') )   or
                       ( ATobRow.GetString('G_GENERAL') = GetParamSocSecur('SO_CPGCDPERTE','') )     or
                       ( ATobRow.GetString('G_GENERAL') = GetParamSocSecur('SO_CPGCDPROVISION','') ) or
                       ( ATobRow.GetString('G_GENERAL') = GetParamSocSecur('SO_CPGCDDOTPROV','') )   or
                       ( ATobRow.GetString('G_GENERAL') = GetParamSocSecur('SO_CPGCDREPRISE','') ) ) ;

    end;
  end;  
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.InitAllPopUp ( vActivation : Boolean ) ;
begin
  if (vActivation) then ChargeFInfoGene;

  InitPopUpComptable    ( vActivation );
  InitPopUpSpecifique   ( vActivation );
  InitPopUpAideRevision ( vActivation );
  InitPopUpRevInteg     ( vActivation );
  InitPopUpParametre    ( vActivation );
  InitPopUpEdition      ( vActivation );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnPopUpComptable(Sender: TObject);
begin
  InitPopUpComptable( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2007
Modifié le ... : 07/02/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.InitPopUpComptable( vActivation : Boolean );
var i : integer;
    lMenuitem : TMenuItem;
begin
  if vActivation then ChargeFInfoGene;

  for i := 0 to PopUpComptable.Items.Count -1 do
  begin
    lMenuItem := PopUpComptable.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'DETAILCOMPTE' then
      begin
        if vActivation then
          lMenuItem.Enabled := (FInfoGene.General <> '')
        else
          lMenuItem.OnClick := OnClickDetailCompte;
        Continue;
      end;

      if lMenuItem.Name = 'SAISIEBOR' then
      begin
        if vActivation then
          lMenuItem.Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures), False)
        else
          lMenuItem.OnClick := SaisieBorClick;
        Continue;
      end;

      if lMenuItem.Name = 'SAISIEPIECE' then
      begin
        if vActivation then
          lMenuItem.Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures), False)
        else
          lMenuItem.OnClick := SaisieEcrClick;
        Continue;
      end;

      if lMenuItem.Name = 'SAISIETRESO' then
      begin
        if vActivation then
          lMenuItem.Enabled:= FInfoGene.SaisieTreso
        else
          lMenuItem.OnClick := SaisieTresoClick;
        Continue;
      end;

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnPopUpSpecifique(Sender: TObject);
begin
  InitPopUpSpecifique( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.InitPopUpSpecifique( vActivation : Boolean );
var i,j       : integer;
    lMenuItem : TMenuItem;
begin
  if vActivation then ChargeFInfoGene;

  for i := 0 to PopUpSpecifique.Items.Count -1 do
  begin
    lMenuItem := PopUpSpecifique.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'TRAITEMENTSPECIFIQUE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'CPCONSAUXI' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Collectif
            else
              lMenuItem.Items[j].OnClick := ConsAuxiliaireClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'LISTEIMMO' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Immo
            else
              lMenuItem.Items[j].OnClick := ListeImmoClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DETAILICC' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Icc
            else
              lMenuItem.Items[j].OnClick := OnClickDetailIcc;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DETAILCAPITALICC' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.IccCapital
            else
              lMenuItem.Items[j].OnClick := OnClickDetailIcc;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'GCD' then
          begin // Créer une créance dans GCD
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.GCD
            else
              lMenuItem.Items[j].OnClick := OnClickGCD ;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'EMPRUNTCRE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Cre
            else
              lMenuItem.Items[j].OnClick := OnClickEmpruntCRE;
            Continue;
          end;
        end; // FOR
      end; // TRAITEMENTSPECIFIQUE

      if lMenuItem.Name = 'LETTRAGE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'LETMANUEL' then
          begin
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Lettrable
            else lMenuItem.Items[j].OnClick := LettrageMClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'LETAUTO' then
          begin
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Lettrable
            else lMenuItem.Items[j].OnClick := LettrageAClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DELETEXOREF' then
          begin // Delettrage compelt pour l'exercice de référence
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Lettrable
            else lMenuItem.Items[j].OnClick := DelettreMvtExoRef;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DELETTOTAL' then
          begin // Delettrage total du compte
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Lettrable
            else lMenuItem.Items[j].OnClick := DelettreTotal;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'PASSEJUSTIFSOLDE' then
          begin // Passage en justif de solde par CPCONSECR
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Lettrable
            else
              lMenuItem.Items[j].OnClick := OnClickPasseJustifSolde;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'ETATJUSTIFSOLDE' then
          begin // Justificatif de solde (ETAT)
            if vActivation then
            lMenuItem.Items[j].Enabled := FInfoGene.Lettrable
            else
            lMenuItem.Items[j].OnClick := OnClickJustifSolde;
            Continue;
          end;
        end; // FOR
      end; // LETTRAGE

      if lMenuItem.Name = 'ANALYTIQUE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'MVTANALYTIQ' then
          begin
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Analytique
            else lMenuItem.Items[j].OnClick := AnalytiquesClick;
            Continue;
          end;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'REPARTANALYTIQ' then
          begin
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Analytique
            else lMenuItem.Items[j].OnClick := RepartitionAnaClick;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          if lMenuItem.Items[j].Name = 'GLANASURAXE1' then
          begin // Grand livre analytique sur l'axe 1
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnaSurAxe1
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe1;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'GLANASURAXE2' then
          begin // Grand livre analytique sur l'axe 2
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnaSurAxe2
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe2;
            Continue;
          end;

          // Grand livre analytique sur l'axe 3
          if lMenuItem.Items[j].Name = 'GLANASURAXE3' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnaSurAxe3
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe3;
            Continue;
          end;

          // Grand livre analytique sur l'axe 4
          if lMenuItem.Items[j].Name = 'GLANASURAXE4' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnaSurAxe4
            else
               lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe4;
            Continue;
          end;

          // Grand livre analytique sur l'axe 5
          if lMenuItem.Items[j].Name = 'GLANASURAXE5' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnaSurAxe5
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe5;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE1' then
          begin //  Balance Analytique sur l'axe 1
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnasurAxe1
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe1;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE2' then
          begin //  Balance Analytique sur l'axe 2
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnasurAxe2
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe2;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE3' then
          begin //  Balance Analytique sur l'axe 3
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnasurAxe3
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe3;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE4' then
          begin //  Balance Analytique sur l'axe 4
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnasurAxe4
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe4;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE5' then
          begin //  Balance Analytique sur l'axe 5
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0) and
                                            FInfoGene.Analytique and
                                            FInfoGene.AnasurAxe5
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe5;
            Continue;
          end;
        end; // FOR
      end; // ANALYTIQUE

      if lMenuItem.Name = 'POINTAGE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'POINTAGEMANUEL' then
          begin
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Pointable
            else lMenuItem.Items[j].OnClick := PointageGClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'POINTAGEENCOURS' then
          begin
            if vActivation then lMenuItem.Items[j].Enabled := FInfoGene.Pointable
            else lMenuItem.Items[j].OnClick := PointageEnCoursClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'ETATRAPPRO' then
          begin // Etat de rapprochement
            if vActivation then
              lMenuItem.Items[j].Enabled := (not VH^.PointageJal) and (FInfoGene.General <> '') and
                                            (FInfoGene.Pointable) and (ATobFListe.Detail.Count > 0)
            else
              lMenuItem.Items[j].OnClick := OnClickEtatRapprochement;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'ETATJUSTIFSOLDEBQE' then
          begin // GCO - 05/07/2007 - Justificatif de Solde Bancaire FQ 20926
            if vActivation then
            {$IFDEF COMPTA}
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '') and
                                                 (not VH^.PointageJal) and
                                                 (FInfoGene.Pointable) and
                                                 (FInfoGene.NatureGene = 'BQE') and
                                                 (ATobFListe.Detail.Count > 0)
            {$ENDIF}
            else
              lMenuItem.Items[j].OnClick := OnClickJustifSoldeBQE;
            Continue;
          end;




        end; // FOR
      end; // POINTAGE

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
     end;
  end; // for
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnPopUpAideRevision(Sender: TObject);
begin
  InitPopUpAideRevision( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.InitPopUpAideRevision( vActivation : Boolean );
var i, j : integer;
    lMenuItem : TMenuItem;
    lStGen : string;
    lTob : Tob;
begin
  if vActivation then ChargeFInfoGene;

  for i := 0 to PopUpAideRevision.Items.Count -1 do
  begin
    lMenuItem := PopUpAideRevision.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'OUTILREVISION' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'NOTETRAVAIL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModRIC and (FInfoGene.General <> '') and JaiLeRoleCompta( RcReviseur )
            else
              lMenuItem.Items[j].OnClick := OnClickNoteTravail;
            Continue;  
          end;

          if lMenuItem.Items[j].Name = 'TABLEAUVARIATION' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModRIC and (FInfoGene.General <> '') and JaiLeRoleCompta( RcReviseur )
            else
             lMenuItem.Items[j].OnClick := OnClickTableauVariation;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'RECHERCHEECR' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := True
            else
              lMenuItem.Items[j].OnClick := OnClickRechercheEcritures;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'FEUILLEEXCEL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModRIC and JaiLeRoleCompta( RcReviseur )
            else
              lMenuItem.Items[j].OnClick := OnClickFeuilleExcel;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'CUTOFF' then
          begin
            if vActivation then
            begin
              lTob := GetO(FListe) ;
              if lTob <> nil then
              begin
                lStGen := Copy(lTob.GetString('G_GENERAL'),1,3) ;
                lMenuItem.Items[j].Enabled := ( lTob.GetString('CUTOFF') = 'X' ) or
                                              ( ( lStGen = '486' ) or ( lStGen = '487' ) ) ;
              end;
            end
            else
              lMenuItem.Items[j].OnClick := OnClickCutOff ;
            Continue;  
          end;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'CONTROLETVA' then
          begin
            if vActivation then
            begin
              lStGen := Copy(FInfoGene.General, 1, 4);
              lMenuItem.Items[j].Enabled := (lStGen = '4456') or (lStGen = '4457');
            end
            else
              lMenuItem.Items[j].OnClick := OnClickControleTva;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          if lMenuItem.Items[j].Name = 'CALCULICC' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Icc or FInfoGene.IccCapital
            else
              lMenuItem.Items[j].OnClick := OnClickCalculIcc;
            Continue;
          end;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'DRF' then
          begin // Détermination du résultat fiscal
            if vActivation then
              lMenuItem.Items[j].Enabled := (FExoDate.Code = VH^.EnCours.Code) and
                                            (VH^.OkModDRF) and
                                            (OkLanceDRF <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickDRF;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          if lMenuItem.Items[j].Name = 'JUSTIFGCD' then
          begin // Tableau Justificatif des créances douteuses
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.GCDCOMPTA
            else
              lMenuItem.Items[j].OnClick := OnClickTableauGCD ;
            Continue;
          end;

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'RAPPROCOMPTACRE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Cre
            else
              lMenuItem.Items[j].OnClick := OnClickRappECRCRE;
            Continue;
          end;
          {$ENDIF}

          if lMenuItem.Items[j].Name = 'RAPPROCOMPTAIMMO' then
          begin
            if vActivation then
            {$IFDEF AMORTISSEMENT}
              lMenuItem.Items[j].Enabled := FInfoGene.Immo
            {$ENDIF}
            else
              lMenuItem.Items[j].OnClick := OnClickRapproComptaImmo;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'RAPPROCOMPTAGCD' then
          begin // Rapprochement Comptabilité / GCD
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.GCDCOMPTA
            else
              lMenuItem.Items[j].OnClick := OnClickRapproComptaGCD ;
            Continue;
          end;

        end; // FOR
      end; // OUTILREVISION

      if lMenuItem.Name = 'LIASSEFISCALE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'INFOLIASSE' then
          begin // Emplacement dans la liasse fiscale
            if vActivation then
              lMenuItem.Items[j].Enabled := (FStLiasseDossier <> '') and
                                            (FInFoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickInfoLiasse;
            Continue;  
          end;

          if lMenuItem.Items[j].Name = 'TOTALCPTLIASSE' then
          begin // Montant affecté dans la liasse fiscale
            if vActivation then
              lMenuItem.Items[j].Enabled := (FStLiasseDossier <> '') and
                                            (FInfoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickTotalCptLiasse;
          end;

          if lMenuItem.Items[j].Name = 'DONNEESCOMP' then
          begin // Données complémentaires
            if vActivation then
              lMenuItem.Items[j].Enabled := (FStLiasseDossier <> '') and
                                            (FInfoGene.General <> '') and CPPresenceEtafi
            else
              lMenuItem.Items[j].OnClick := OnClickDonneesComp;
            Continue;  
          end;

        end; // FOR
      end; // LIASSEFISCALE

      if lMenuItem.Name = 'ANALYSE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'CUMULCPTE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := CumulsGENEClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'HISTOCPTE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := HistoCpteClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'ANALYSCOMP' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := ANALREVISClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'MEMOMILLESIME1' then
          begin // Mémo commantire millésimé
            if vActivation then // GCO - 04/05/2007 - FQ 20244
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '') and (not VH^.OkModRIC)
            else
              lMenuItem.Items[j].OnClick := OnClickMemoMillesime;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'MEMOCOMPTE1' then
          begin // Mémo compte général
            if vActivation then // GCO - 04/05/2007 - FQ 20244
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '') and (not VH^.OkModRIC)
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCompte;
            Continue;
          end;
        end; // FOR
      end; // ANALYSE

      if lMenuItem.Name = 'VISACOMPTE' then
      begin
        if vActivation then
        begin
          if (FInfoGene.General <> '') then
          begin
            if FInfoGene.VisaRevision then
              lMenuItem.Caption := TraduireMemoire('Enlever le visa du compte')
            else
              lMenuItem.Caption := TraduireMemoire('Viser le compte');
          end;
          lMenuItem.Enabled := (FInfoGene.General <> '') and
                               (FExoDate.Code = VH^.Encours.Code) and
                               AutoriseSuppresionVisaRevision(FInfoGene.General);
        end
        else
          lMenuItem.OnClick := OnClickVisaCompte;
        Continue;
      end;

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnPopUpRevInteg(Sender: TObject);
begin
  InitPopUpRevInteg( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.InitPopUpRevInteg( vActivation : Boolean );
var i,j : integer;
    lMenuItem : TMenuItem;
begin
  if vActivation then ChargeFInfoGene;

  for i := 0 to PopUpRevInteg.Items.Count -1 do
  begin
    lMenuItem := PopUpRevInteg.Items[i];
    try
      lMenuItem.Enabled := False;

      {$IFDEF COMPTA}
      {$IFNDEF CCMP}
      if lMenuItem.Name = 'PROGTRAVAIL' then
      begin
        // GCO - 25/07/2007 - FQ 20944
        lMenuItem.Caption := TraduireMemoire('Programme de &travail');

        if vActivation then
          lMenuItem.Enabled := VH^.OkModRIC and FBoOkRevision and
                               ((JaiLeRoleCompta(rcSuperviseur)) or
                                (JaiLeRoleCompta(rcReviseur) and
                                 not JaileRoleCompta(rcSuperviseur))
                               )
        else
          lMenuItem.OnClick := OnClickProgrammeTravail;
        Continue;
      end;
      {$ENDIF}
      {$ENDIF}

      if lMenuItem.Name = 'DOCTRAVAUX' then
      begin // Documentation des travaux dépend de la séria RIC
        lMenuItem.Enabled := VH^.OkModRIC;

        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'MEMOCYCLE' then
          begin // Mémo cycle
            if vActivation then
              lMenuItem.Items[j].Enabled := JaiLeRoleCompta( RcReviseur ) and
                                            FBoOkRevision and (FInfoGene.CycleRevision <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCycle;
          end;

          if lMenuItem.Items[j].Name = 'MEMOOBJECTIF' then
          begin // Mémo objectif de révision
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl In V_Pgi.PgiContexte) and
                                            JaiLeRoleCompta( RcReviseur ) and
                                            FBoOkRevision and (FInfoGene.CycleRevision <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoObjectif;
          end;

          if lMenuItem.Items[j].Name = 'MEMOSYNTHESE' then
          begin // Mémo synthèse du cycle
            if vActivation then
              lMenuItem.Items[j].Enabled := JaiLeRoleCompta( RcReviseur ) and
                                            FBoOkRevision and (FInfoGene.CycleRevision <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoSynthese;
          end;

          if lMenuItem.Items[j].Name = 'MEMOMILLESIME2' then
          begin // Mémo commantire millésimé
            if vActivation then
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoMillesime;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'MEMOCOMPTE2' then
          begin // Mémo compte général
            if vActivation then
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCompte;
            Continue;
          end;
        end; // FOR
      end; // DOCTRAVAUX

      {$IFDEF COMPTA}
      {$IFNDEF CCMP}
      if lMenuItem.Name = 'SUPERVISIONTRAVAUX' then
      begin
        // GCO - 11/06/2007 - FQ 20584 + FQ 21365
        lMenuItem.Enabled := (CtxPcl In V_Pgi.PgiContexte) and VH^.OkModRIC and
                             FBoOkRevision and JaiLeRoleCompta( rcSuperviseur );
                             
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := lMenuItem.Enabled; // GCO - 13/09/2007 - FQ 2136
          if lMenuItem.Items[j].Name = 'APG' then
          begin
            if not vActivation then // Appréciation générale du dossier
              lMenuItem.Items[j].OnClick := OnClickAPG;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'SCY' then
          begin
            if not vActivation then // ynthèse des cycles
              lMenuItem.Items[j].OnClick := OnClickSCY;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'EXP' then
          begin
            if not vActivation then // Attestation Expert
              lMenuItem.Items[j].OnClick := OnClickEXP;
            Continue;
          end;
        end;
      end;
      {$ENDIF}
      {$ENDIF}

      if lMenuItem.Name = 'DOSSIERCLIENT' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'DOSPERMANENT' then
          begin // Dossier permanent
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickDocPermanent;
          end;

          if lMenuItem.Items[j].Name = 'DOSANNUEL' then
          begin // Dossier annuel
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickDocAnnuel;
          end;

          if lMenuItem.Items[j].Name = 'RECHERCHEDOC' then
          begin // Recherche de document
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickRechercheDoc;
          end;
        end; // FOR
      end; // DOSSIERCLIENT

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnPopUpParametre(Sender: TObject);
begin
  InitPopUpParametre( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2007
Modifié le ... : 07/02/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.InitPopUpParametre( vActivation : Boolean );
var i,j : integer;
    lMenuItem : TMenuItem;
begin
  if vActivation then ChargeFInfoGene;

  for i := 0 to PopUpParametre.Items.Count - 1 do
  begin
    lMenuItem := PopUpParametre.Items[i];
    try
      lMenuItem.Enabled := False;

      if lMenuItem.Name = 'PARAMETRES' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'PARAMGENERAL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FInfoGene.General <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickCompteGeneral;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'RIBBQE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FInfoGene.Bqe
            else
              lMenuItem.Items[j].OnClick := OnClickCompteRIB;
            Continue;
          end;
            end; // FOR
      end; // PARAMETRES

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnPopUpEdition(Sender: TObject);
begin
  InitPopUpEdition( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.InitPopUpEdition(vActivation: Boolean);
var i : integer;
    lMenuItem : TMenuItem;
begin
  if vActivation then ChargeFInfoGene;

  for i := 0 to PopUpEdition.Items.Count - 1 do
  begin
    lMenuItem := PopUpEdition.Items[i];
    try
      lMenuItem.Enabled := False;

      // GCO - 23/10/2007 - FQ 21677
      if lMenuItem.Name = 'EJUSTIFSOLDEBQE' then
      begin // Justificatif de Solde Bancaire PGE
        if vActivation then
        {$IFDEF COMPTA}
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (FInfoGene.General <> '') and
                               (not VH^.PointageJal) and
                               (FInfoGene.Pointable) and
                               (FInfoGene.NatureGene = 'BQE') and
                               (ATobFListe.Detail.Count > 0)
        {$ENDIF}
        else
          lMenuItem.OnClick := OnClickJustifSoldeBQE;
        Continue;
      end;

      if lMenuItem.Name = 'EETATRAPPRO' then
      begin // Etat de rapprochement PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (not VH^.PointageJal) and (FInfoGene.General <> '') and
                               (FInfoGene.Pointable) and (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickEtatRapprochement;
        Continue;
      end;

      if lMenuItem.Name = 'EETATJUSTIFSOLDE' then
      begin // Justificatif de solde PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and FInfoGene.Lettrable
        else
          lMenuItem.OnClick := OnClickJustifSolde;
        Continue;
      end;

      if lMenuItem.Name = 'GLGENE' then
      begin // Grand Livre Général
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickGLGene;
        Continue;
      end;

      if lMenuItem.Name = 'GLGENEPARQTE' then
      begin // Grand Livre Général par Quantité
        if vActivation then
          lMenuItem.Enabled := FBoOkGereQte and (FInfoGene.General <> '') and
                               ((FInfoGene.QualifQte1 <> '') or (FInfoGene.QualifQte2 <> ''))
        else
          lMenuItem.OnClick := OnClickGLGeneParQte;
        Continue;
      end;

      if lMenuItem.Name = 'GLGENEPARAUXI' then
      begin // Grand livre général par auxiliaire
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0) and FInfoGene.Collectif
        else
          lMenuItem.OnClick := OnClickGLGeneParAuxi;
        Continue;
      end;

      // Grand livre analytique sur l'axe 1 PGE
      if lMenuItem.Name = 'EGLANASURAXE1' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnaSurAxe1
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe1;
        Continue;
      end;

      // Grand livre analytique sur l'axe 2 PGE
      if lMenuItem.Name = 'EGLANASURAXE2' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnaSurAxe2
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe2;
        Continue;
      end;

      // Grand livre analytique sur l'axe 3 PGE
      if lMenuItem.Name = 'EGLANASURAXE3' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnaSurAxe3
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe3;
        Continue;
      end;

      // Grand livre analytique sur l'axe 4 PGE
      if lMenuItem.Name = 'EGLANASURAXE4' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnaSurAxe4
        else
           lMenuItem.OnClick := OnClickGLAnaSurAxe4;
        Continue;
      end;

      // Grand livre analytique sur l'axe 5 PGE
      if lMenuItem.Name = 'EGLANASURAXE5' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnaSurAxe5
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe5;
        Continue;
      end;

      if lMenuItem.Name = 'BALGENE' then
      begin // Balance générale
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickBalGene ;
        Continue;
      end;

      // Balance générale par auxiliaire
      if lMenuItem.Name = 'BALGENEPARAUXI' then
      begin
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0) and FInfoGene.Collectif
        else
          lMenuItem.OnClick := OnClickBalGeneParAuxi;
        Continue;
      end;

      //  Balance Analytique sur l'axe 1 PGE
      if lMenuItem.Name = 'EBALANASURAXE1' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnasurAxe1
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe1;
        Continue;
      end;

      // Balance Analytique sur l'axe 2 PGE
      if lMenuItem.Name = 'EBALANASURAXE2' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnasurAxe2
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe2;
        Continue;
      end;

      //  Balance Analytique sur l'axe 3 PGE
      if lMenuItem.Name = 'EBALANASURAXE3' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnasurAxe3
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe3;
        Continue;
      end;

      if lMenuItem.Name = 'EBALANASURAXE4' then
      begin //  Balance Analytique sur l'axe 4
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnasurAxe4
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe4;
        Continue;
      end;

      if lMenuItem.Name = 'EBALANASURAXE5' then
      begin //  Balance Analytique sur l'axe 5
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FInfoGene.Analytique and
                               FInfoGene.AnasurAxe5
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe5;
        Continue;
      end;

      if lMenuItem.Name = 'NOTECTRLCOMPTE' then
      begin // Note de Contrôle Compte
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickNoteCtrlCompte;
        Continue;
      end;

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/11/2002
Modifié le ... :   /  /
Description .. : Affiche la fenêtre des paramètres du compte général
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickCompteGeneral(Sender: TObject);
var lTob : Tob;
    lAction : TActionFiche ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  lAction := TaModif;
  FicheGene(nil,'',lTob.GetString('G_GENERAL'), lAction, 0);
  BCherche.Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickDonneesComp(Sender: TObject);
var
  lObj : Variant;
  lstMsg : string;
  lstTableau : string;
  lTob : Tob;
  lstGeneral : string;
  i : integer;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  lstGeneral :=  lTob.GetString('G_GENERAL');
  if (Copy(lstGeneral,1,2)='70') then lstTableau := 'SCVTILCA'
  else if ((lTob.GetString('NATUREGENE')='IMO') and (Copy(lstGeneral,1,2)<>'28')
    and (Copy(lstGeneral,1,2)<>'29')) then lstTableau := 'SC2054P1'
  else if (lTob.GetString('NATUREGENE')='IMO') then lstTableau := 'SC2055P1'
  else if (Copy(lstGeneral,1,3)='145') then lstTableau := 'SC2055P2'
  else if (Copy(lstGeneral,1,1)='4') then lstTableau := 'SC2057P2'
  else if ((Copy(lstGeneral,1,2)='67') or (Copy(lstGeneral,1,2)='77')) then lstTableau := 'SC2053P1'
  else lstTableau := 'SCVTILCA';

  lObj := CreateOleObject('CegidPgi.CegidPgi');
  lstMsg := lObj.LanceEtafi (V_PGI.NoDossier);
  if lstMsg <> '' then
  begin
    Application.BringToFront;
    PGIBox(lstMsg);
    lObj := unassigned;
  end else
  begin
    lObj := unassigned;
    // Temporisation nécessaire pour attendre que Etafi soit correctement lancé
    for i := 0 to 20 do
    begin
      if ExisteSQL ('SELECT 1 FROM COURRIER WHERE MG_TYPE=42 AND MG_EXPEDITEUR="'+V_PGI.User+'"') then
      begin
        lObj := CreateOleObject('CegidPgi.LiasseBIC');
        lstMsg := lObj.SaisieComplementaireEtafi('E', VH^.Encours.Code, lstTableau);
        if lstMsg <> '' then
        begin
          Application.BringToFront;
          PGIBox(lstMsg);
        end;
        lObj := unassigned;
        break;
      end;
      sleep (1000);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickInfoLiasse(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CPLancefiche_CPControleLiasse(lTob.GetString('G_GENERAL'));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickTotalCptLiasse(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CPLancefiche_CPTotalCptLiasse(lTob.GetString('G_GENERAL') + ';' +
                                DateToStr(FExoDate.Deb) + ';' +
                                DateToStr(FExoDate.Fin));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickCompteRIB(Sender: TObject);
var lTob : Tob;
    lOkExiste : Boolean;
    lAction : TActionFiche ;

    lStNatureGene : string;
    lStGeneral    : string;

begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  lStNatureGene := lTob.GetString('NATUREGENE');
  lStGeneral    := lTob.Getstring('G_GENERAL');

  // GCO - 25/11/2004 - FQ 14928
  if ( lStNatureGene = 'BQE') or
     ( lStNatureGene = 'TIC') or
     ( lStNatureGene = 'TID') then
  begin
    if lStNatureGene = 'BQE' then
    begin // 24/10/2006 YMO Multisociétés
      lOkExiste := ExisteSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_GENERAL = "' + lStGeneral +
                             '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');
      if lOkExiste then
        lAction := TaModif
      else
        lAction := TaCreatOne;
      FicheBanqueCP(lStGeneral, lAction, 0) ;
    end
    else
    begin // Compte  TIC ou TID
      lOkExiste := ExisteSQL('SELECT R_AUXILIAIRE FROM RIB WHERE R_AUXILIAIRE = "' + lStGeneral + '"');
      if lOkExiste then
        lAction := TaModif
      else
        lAction := TaCreatOne;
      FicheRIB_AGL( lStGeneral, lAction, False, '', False);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/02/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickVisaCompte(Sender: TObject);
var lTob : Tob;
    lBoVisaRevision : Boolean;
begin
  // GCO - 23/08/2005 - FQ 16248
  if (FExoDate.Code = VH^.Encours.Code) then
  // and (not ExisteSql('SELECT MG_EXPEDITEUR FROM COURRIER')) then
  begin
    lTob := GetO(FListe) ;
    if (lTob = nil) then Exit;
    lBoVisaRevision := IIF( lTob.GetString('G_VISAREVISION') = 'X', False, True);

    if not MiseAJourG_VisaRevision( lTob.GetString('G_GENERAL'), lBoVisaRevision )  then
      PgiError('Impossible de mettre à jour les informations de révision du compte.', Ecran.Caption);

    RefreshPclPGE;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickDocPermanent(Sender: TObject);
var
  lObj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Dossier permanent');
  lObj := unassigned;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickDocAnnuel(Sender: TObject);
Var
  lobj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Dossier annuel');
  lObj := unassigned;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickRechercheDoc(Sender: TObject);
var
  lObj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Recherche documentaire');
  lObj := unassigned;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... : 11/04/2007
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CPCONSGENE.OnClickProgrammeTravail(Sender: TObject);
var lStRacine : HString;
begin
  // GCO - 18/09/2007 - FQ 21451
  lStRacine := '';
  if G_CycleRevision.Text <> '' then
    lStRacine := Copy(G_CycleRevision.Text, 0, 1);

  CPLanceFiche_CPRevProgTravail( VH^.EnCours.Code, lStRacine );
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickMemoCycle(Sender: TObject);
begin
{$IFDEF COMPTA}
  CPLanceFiche_CPRevDocTravaux( FInfoGene.General, '', VH^.EnCours.Code, 0 );
  RefreshPclPGE;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickMemoObjectif(Sender: TObject);
begin
{$IFDEF COMPTA}
  CPLanceFiche_CPREVDOCTRAVAUX( FInfoGene.General, '', VH^.EnCours.Code, 1 );
  RefreshPclPGE;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickMemoSynthese(Sender: TObject);
begin
{$IFDEF COMPTA}
  CPLanceFiche_CPREVDOCTRAVAUX( FInfoGene.General, '', VH^.EnCours.Code, 2 );
  RefreshPclPGE;
{$ENDIF}  
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickMemoMillesime(Sender: TObject);
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if (lTob = nil) or (FExoDate.Code <> VH^.Encours.Code) then Exit;
  CPLanceFiche_CPREVDOCTRAVAUX( ATobRow.GetString('G_GENERAL'), '', VH^.EnCours.Code, 3 );
  RefreshPclPGE;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickMemoCompte(Sender: TObject);
begin
{$IFDEF COMPTA}
  CPLanceFiche_CPREVDOCTRAVAUX( ATobRow.GetString('G_GENERAL'), '', VH^.EnCours.Code, 4 );
  RefreshPclPGE;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CPCONSGENE.OnClickAPG( Sender : TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 1);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CPCONSGENE.OnClickSCY( Sender : TObject );
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 2);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CPCONSGENE.OnClickEXP( Sender : TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 3);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBBlocNote(Sender: TObject);
var lTobGene : Tob;
begin
  inherited;
  HPB.Visible := BBlocNote.Down;
  if BBlocNote.Down then
  begin
    FBLocNote.Clear;
    lTobGene := GetO(FListe) ;
    if lTobGene = nil then Exit;
    // GCO - 31/05/2007 - FQ 20479
    StringToRich( FBlocNote, ChargeMemoCREVBLOCNOTE('GEN', lTobGene.GetString('G_GENERAL'), VH^.Encours.Code));
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBCyclePrecedent(Sender: TObject);
var lSt : string;
begin
  lSt := G_CycleRevision.Text;
  lSt := ReadTokenSt( lSt );
  lSt := GetColonneSql('GENERAUX', 'G_CYCLEREVISION', 'G_CYCLEREVISION < "' + lSt + '" ORDER BY G_CYCLEREVISION DESC');
  if lSt <> '' then
  begin
    G_CycleRevision.Text := lSt;
    BCherche.Click;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBCycleSuivant(Sender: TObject);
var lSt : string;
begin
  lSt := G_CycleRevision.Text;
  lSt := ReadTokenSt( lSt );
  lSt := GetColonneSql('GENERAUX', 'G_CYCLEREVISION', 'G_CYCLEREVISION > "' + lSt + '" ORDER BY G_CYCLEREVISION ASC');
  if lSt <> '' then
  begin
    G_CycleRevision.Text := lSt;
    BCherche.Click;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnCkickCBCptSansCycle( Sender : TObject );
begin
  // GCO - 15/05/2007 - FQ 18680
  G_CYCLEREVISION.Enabled := (CBCptSansCycle.State = CbGrayed) or
                             (CBCptSansCycle.State = CbUnChecked);

  if not G_CycleREvision.Enabled then
    G_CycleRevision.Text := '';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : __/__/____
Modifié le ... : 05/01/2007
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnPopUpPopF11(Sender: TObject);
begin
  PurgePopup( PopF11 );
  AddMenuPop( PopF11, '', '');
  InitAllPopUp( True );
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/11/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
  case Key of

    VK_F3 : if FBoOkRevision then
            begin
              Key := 0;
              OnClickBCyclePrecedent( nil );
            end;

    VK_F4 : if FBoOkRevision then
            begin
              Key := 0;
              OnClickBCycleSuivant( nil );
            end;  

    VK_F5 : if FListe.Focused then
            begin
              Key := 0; OnClickDetailCompte( nil );
            end ;


    Ord('C') : begin
                 if ssAlt in Shift then
                 begin // ALT + C
                   Key := 0;
                   OnClickCommentCpte( nil );
                 end;
               end;

    // GCO - 14/06/2006 - FQ 17669
    Ord('G') : begin // CTRL + G -> Accès zone GENERAL
                 if ssCtrl in Shift then
                 begin
                   PageControl.ActivePage := PageControl.Pages[0];
                   if G_General.CanFocus then
                     SetFocusControl('G_GENERAL');
                 end;
               end;

    // ALT + O  -> Commnentaire millésimé
    Ord('O') : if ssAlt in Shift then
               begin
                 Key := 0;
                 OnClickMemoMillesime(nil);
               end ;

    Ord('R') : begin
                 if ssAlt in Shift then
                 begin // ALT + R -> accès au RIB
                   Key := 0;
                   OnClickCompteRib( nil );
                 end;

                 // GCO - 25/07/2007 - FQ 20827
                 if ssCtrl in Shift then // Ctrl + R -> Onglet Révision
                 begin
                   PageControl.ActivePage := TabRevision;
                   G_CycleRevision.SetFocus;
                   G_CycleREvision.SelStart := 0;
                   G_CycleREvision.SelLength := Length(G_CycleREvision.Text);
                 end;
               end;

    Ord('V') : begin
                 if ssAlt in Shift then
                 begin // ALT + V -> Visa du compte
                   Key := 0;
                   // GCO - 10/05/2007 - FQ 20273
                   ChargeFInfoGene;
                   if (FInfoGene.General <> '') and
                      (FExoDate.Code = VH^.Encours.Code) and
                      AutoriseSuppresionVisaRevision(FInfoGene.General) then
                     OnClickVisaCompte( nil );
                 end;
               end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Lance la consultation des écritures du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickDetailCompte(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  OperationsSurComptes(ATobRow.GetString('G_GENERAL'), '', '', '',false, '', False, GModeSelection.Value);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSGENE.BeforeLoad : Boolean;
var i : integer;
    lStWhereCompte     : string;
    lStWhereExclusion  : string;
    lStWhereActivation : string;
    lQuery : Tquery;

    lSt  : string;
    lSt1 : string;
    lSt2 : string;

    lTobTemp : Tob;

begin
  FBoCriterePonderation := (GTypePonderation.ItemIndex <> -1) and
                           (ZCPonderation.ItemIndex <> -1);
                           // GCO - 07/06/2007 - FQ 20613
                           //and (GValeurPonderation.Text <> '');

  FBoCritereVariation   := (GTypeVariation.ItemIndex <> -1) and
                           (ZCVariation.ItemIndex <> -1) and
                           // GCO - 07/06/2007 - FQ 20613
                           //(GValeurVariation.Text <> '') and
                           (GNatureVariation.ItemIndex <> -1);

  FStPrefixeE := IIF(FExoDate.Code = VH^.Encours.Code, 'E', 'S');
  FStPrefixeP := IIF(FExoDate.Code = VH^.Encours.Code, 'P', 'E');

  FStCodeExoE := IIF(FExoDate.Code = VH^.Encours.Code, VH^.Encours.Code, VH^.Suivant.Code);
  FStCodeExoP := IIF(FExoDate.Code = VH^.Encours.Code, VH^.Precedent.Code, VH^.Encours.Code);

  if FBoOkRevision then
  begin
    FTobListeCycle.LoadDetailFromSQL('SELECT CCY_CODECYCLE, CCY_ETATCYCLE, ' +
                  'CPC_LISTECOMPTE, CPC_LISTEEXCLUSION, CPC_LISTEACTIVE, ' +
                  '0 TOTAL FROM CREVCYCLE ' +
                  'LEFT JOIN CREVPARAMCYCLE ON CCY_CODECYCLE = CPC_CODECYCLE ' +
                  'WHERE CCY_ACTIVECYCLE = "X" ' +
                  'ORDER BY CCY_CODECYCLE', False);

    for i := 0 to FTobListeCycle.Detail.Count - 1 do
    begin
      lTobTemp := FTobListeCycle.Detail[i];

      lStWhereCompte     := lTobTemp.GetString('CPC_LISTECOMPTE');
      lStWhereExclusion  := lTobTemp.GetString('CPC_LISTEEXCLUSION');
      lStWhereActivation := lTobTemp.GetString('CPC_LISTEACTIVE');

      lSt1 := '';
      lSt2 := '';
      while (lStWhereActivation <> '') do
      begin
        lSt := ReadTokenSt( lStWhereActivation );
        if lSt = 'X' then
        begin
          lSt1 := lSt1 + ReadTokenSt( lStWhereCompte ) + ';' ;
          lSt2 := lSt2 + ReadTokenSt( lStWhereExclusion ) + ';' ;
        end
        else
        begin
          ReadTokenSt( lStWhereCompte );
          ReadTokenSt( lStWhereExclusion );
        end;
      end;

      lStWhereCompte    := AnalyseCompte( lSt1, fbGene, False, False) ;
      lStWhereExclusion := AnalyseCompte( lSt2, fbGene, True, False);

      if lStWhereCompte <> '' then
      begin
        lQuery := OpenSQL('SELECT SUM(G_TOTDEBE - G_TOTCREE) TOTAL FROM GENERAUX WHERE ' +
                          lStWhereCompte +
                          IIF( lStWhereExclusion <> '', ' AND ' + lStWhereExclusion, ''), True);

        lTobTemp.PutValue('TOTAL', lQuery.FindField('TOTAL').asFloat );
        Ferme( lQuery );
      end;

    end;
  end;

  // GCO - 18/06/2007 - Demande KMPG, changer SOLDEE et SOLDEP par le solde
  // d'une balance de situation
  if FTobBDSE <> nil then
  begin
    FTobBDSE.AddchampSupValeur('DATEDEB', '01/01/1900', False);
    FTobBDSE.AddchampSupValeur('DATEFIN', '01/01/1900', False);

    FTobBDSE.LoadDetailFromSQL('SELECT BSE_COMPTE1, SUM(BSE_DEBIT) DE, SUM(BSE_CREDIT) CE, ' +
      'SUM(BSE_DEBIT)-SUM(BSE_CREDIT) TOTAL FROM CBALSITECR WHERE ' +
      'BSE_CODEBAL = "' + FTobBDSE.GetString('BSI_CODEBAL') + '" GROUP BY BSE_COMPTE1');
  end;

  if FTobBDSP <> nil then
  begin
    FTobBDSP.LoadDetailFromSQL('SELECT BSE_COMPTE1, SUM(BSE_DEBIT) DP, SUM(BSE_CREDIT) CP, ' +
      'SUM(BSE_DEBIT)-SUM(BSE_CREDIT) TOTAL FROM CBALSITECR WHERE ' +
      'BSE_CODEBAL = "' + FTobBDSP.GetString('BSI_CODEBAL') + '" GROUP BY BSE_COMPTE1');
  end;    

  Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSGENE.AjouteATobFListe(vTob : Tob) : Boolean;
var lRdSoldeE       : double ;
    lRdSoldeP       : double ;
    lTempSoldeE     : double ;
    lTempSoldeP     : double ;
    lVariation      : double ;
    lPourcentageVariation : double ;
    lStPourcentageVariation : string ;

    lNbMoisSoldeE   : word   ;
    lNbMoisSoldeP   : word   ;
    lPremMois       : word   ;
    lPremAnnee      : word   ;

    lTotalDebit     : double ;
    lTotalCredit    : double ;
    lSoldeAbsolu    : double ;

    lTobBDSE        : Tob;
    lTobBDSP        : Tob;

    function _CalculRevision : Boolean;
    var lBoResultPonderation : Boolean;
        lBoResultVariation   : Boolean;
    begin
      lBoResultPonderation := False;
      lBoResultVariation   := False;

      if FBoCriterePonderation then
      begin
        lBoResultPonderation := False;

        // Type de la pondération
        if GTypePonderation.Value = 'SOLDEDEB' then
        begin
          if lRdSoldeE > 0 then
            lBoResultPonderation := CompareOperateur(lRdSoldeE, FValeurPonderation, ZCPonderation.Value);
        end
        else
        if GTypePonderation.Value = 'SOLDECRE' then
        begin
          if lRdSoldeE < 0 then
            lBoResultPonderation := CompareOperateur(Abs(lRdSoldeE), FValeurPonderation, ZCPonderation.Value);
        end
        else
        if GTypePonderation.Value = 'SOLDE' then
        begin
          lBoResultPonderation := CompareOperateur(lSoldeAbsolu, FValeurPonderation, ZCPonderation.Value);
        end
        else
        if GTypePonderation.Value = 'TOTALDEB' then
        begin
          lBoResultPonderation := CompareOperateur(lTotalDebit, FValeurPonderation, ZCPonderation.Value);
        end
        else
        if GTypePonderation.Value = 'TOTALCRE' then
        begin
          lBoResultPonderation := CompareOperateur(lTotalCredit, FValeurPonderation, ZCPonderation.Value);
        end;
      end;

      if FBoCritereVariation then
      begin
        lBoResultVariation := False;

        if GTypeVariation.Value = 'ABSOLUE' then
        begin
          if GNatureVariation.Value = 'POURCENTAGE' then
          begin
            if (lRdSoldeP <> 0) then // GCO - 11/10/2007 - FQ 21578 Ajout du *100
              lBoResultVariation := CompareOperateur( lPourcentageVariation, FValeurVariation, ZCVariation.Value);
          end
          else
            lBoResultVariation := CompareOperateur( lVariation, FValeurVariation, ZCVariation.Value);
        end
        else
        if GTypeVariation.Value = 'RELATIVE' then
        begin
          if GNatureVariation.Value = 'POURCENTAGE' then
          begin
            if (lRdSoldeP <> 0) then // GCO - 11/10/2007 - FQ 21578 Ajout du *100
              lBoResultVariation := CompareOperateur( lPourcentageVariation, FValeurVariation, ZCVariation.Value);
          end
          else
            lBoResultVariation := CompareOperateur( lVariation, FValeurVariation, ZCVariation.Value);
        end;
      end;

      if FBoCriterePonderation and FBoCritereVariation then
      begin
        Result := lBoResultPonderation and lBoResultVariation
      end
      else
      begin
        if FBoCriterePonderation then
          Result := lBoResultPonderation
        else
          Result := lBoResultVariation;
      end;
    end;
begin

  if FTobBDSE <> nil then
    lTobBDSE := FTobBDSE.FindFirst(['BSE_COMPTE1'], [vTob.GetString('G_GENERAL')], False)
  else
    lTobBDSE := nil;  

  if FTobBDSP <> nil then
    lTobBDSP := FTobBDSP.FindFirst(['BSE_COMPTE1'], [vTob.GetString('G_GENERAL')], False)
  else
    lTobBDSP := nil;


  if FExoDate.Code = VH^.Suivant.Code then
  begin
    // Colonne de G_TOTDEBE et G_TOTCREE sont cachées
    if FColDebitE  <> -1 then FListe.ColWidths[FColDebitE]  := -1;
    if FColCreditE <> -1 then FListe.ColWidths[FColCreditE] := -1;

    // Nombre de mois des exercices pour le calcul de la variation
    NombreMois( VH^.Suivant.Deb, VH^.Suivant.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);

    if lTobBDSE <> nil then
    begin
      vTob.PutValue('G_TOTDEBS', lTobBDSE.GetDouble('DE'));
      vTob.PutValue('G_TOTCRES', lTobBDSE.GetDouble('CE'));
    end;

    if lTobBDSP <> nil then
    begin
      vTob.PutValue('G_TOTDEBE', lTobBDSP.GetDouble('DP'));
      vTob.PutValue('G_TOTCREE', lTobBDSP.GetDouble('CP'));
    end;

    lRdSoldeE := vTob.GetDouble('G_TOTDEBS') - vTob.GetDouble('G_TOTCRES');
    lRdSoldeP := vTob.GetDouble('G_TOTDEBE') - vTob.GetDouble('G_TOTCREE');

    lTotalDebit   := vTob.GetDouble('G_TOTDEBS');
    lTotalCredit  := vTob.GetDouble('G_TOTCRES');
    lSoldeAbsolu  := Abs( lTotalDebit - lTotalCredit);
  end
  else
  begin
    // Colonne de G_TOTDEBS et G_TOTCRES sont cachées
    if FColDebitS  <> -1 then FListe.ColWidths[FColDebitS]  := -1;
    if FColCreditS <> -1 then FListe.ColWidths[FColCreditS] := -1;

    // Nombre de mois des exercices pour le calcul de la variation
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Precedent.Deb, VH^.Precedent.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);

    if lTobBDSE <> nil then
    begin
      vTob.PutValue('G_TOTDEBE', lTobBDSE.GetDouble('DE'));
      vTob.PutValue('G_TOTCREE', lTobBDSE.GetDouble('CE'));
    end;

    if lTobBDSP <> nil then
    begin
      vTob.PutValue('G_TOTDEBP', lTobBDSP.GetDouble('DP'));
      vTob.PutValue('G_TOTCREP', lTobBDSP.GetDouble('CP'));
    end;

    lRdSoldeE := vTob.GetDouble('G_TOTDEBE') - vTob.GetDouble('G_TOTCREE');
    lRdSoldeP := vTob.GetDouble('G_TOTDEBP') - vTob.GetDouble('G_TOTCREP');

    lTotalDebit  := vTob.GetDouble('G_TOTDEBE');
    lTotalCredit := vTob.GetDouble('G_TOTCREE');
    lSoldeAbsolu := Abs( lTotalDebit - lTotalCredit );
  end;

  // Calcul de la Variation
  lTempSoldeE := Abs(lRdSoldeE);
  lTempSoldeP := Abs(lRdSoldeP);

  if (vTob.GetString('SENS') = 'D') or (vTob.GetString('SENS') = 'M') then
  begin
    if (lRdSoldeE < 0) then lTempSoldeE := -1 * lTempSoldeE;
    if (lRdSoldeP < 0) then lTempSoldeP := -1 * lTempSoldeP;
  end
  else
  begin
    if (lRdSoldeE > 0) then lTempSoldeE := -1 * lTempSoldeE;
    if (lRdSoldeP > 0) then lTempSoldeP := -1 * lTempSoldeP;
  end;

  // GCO - 24/05/2006 - FQ 18197
  if GTypeVariation.Value = 'ABSOLUE' then
    lVariation := (lTempSoldeE - lTempSoldeP)
  else
    lVariation := ((lTempSoldeE / lNbMoisSoldeE) * 12) - ((lTempSoldeP / lNbMoisSoldeP) * 12);

  // Calcul du Pourcentage
  if (lRdSoldeP = 0) and (lRdSoldeE <> 0) then // Division par 0
    lPourcentageVariation := cNS
  else
  begin
    if (lRdSoldeP = 0) and (lRdSoldeE = 0) then
      lPourcentageVariation := 0
    else
    begin
      lPourcentageVariation := ((lVariation) / Abs(lTempSoldeP))* 100;
      lPourcentageVariation := Arrondi(lPourcentageVariation, V_PGI.OkDecV);
    end;
  end;

  if Abs(lPourcentageVariation) > 300 then
    lStPourcentageVariation := 'NS'
  else
    lStPourcentageVariation := StrFMontant( lPourcentageVariation, 15, V_PGI.OkDecV, '', True) + ' %';

  if FBoCriterePonderation or FBoCritereVariation then
    Result := _CalculRevision
  else
    Result := True;

  if Result then
  begin
    vTob.AddChampSupValeur('SOLDEE', lRdSoldeE, False);
    vTob.AddChampSupValeur('SOLDEP', lRdSoldeP, False);
    vTob.AddChampSupValeur('VARIATION',   lVariation,   False);
    vTob.AddChampSupValeur('POURCENTAGE', lStPourcentageVariation, False);
    vTob.AddChampSupValeur('XPOURCENTAGE', lPourcentageVariation, False);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2002
Modifié le ... : 26/01/2007
Description .. : Charge ATOBFLISTE dans la grille avec le PutGridDetail
Suite ........ : Retaille la largeur des colonnes de FLISTE
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.RefreshFListe;
begin
  inherited;
  MiseAJourCaptionEcran;

  // GCO - 26/01/2007
  

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var lTob : Tob;
begin
  inherited;
  if ARow = 0 then Exit;
  if (FListe.Row <> ARow)then
  begin
    lTob := GetO(FListe, ARow) ;
    if lTob <> nil then
    begin
      if (ACol = FColSoldeP) or (ACol = FColSoldeE) then
      begin
        if lTob.GetString('SENS') = 'D' then
          Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed, ClGreen)
        else
          if lTob.GetString('SENS') = 'C' then
            Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClGreen, ClRed);
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var lTobCycle : Tob;
    lStCycle  : string;
    lDbTotalCompte : double;
    lDbTotalCycle  : double;
    lDbTotalPonderation : double;
begin
  inherited;

  if ATobRow = nil then Exit;

  OnClickBBlocNote(nil);

  if not FBoOkRevision then Exit;

  AssigneImage( ImNoteTravail , ATobRow.GetString('NOTETRAVAIL') = 'X', 1);
  AssigneImage( IMTableauVariation, ATobRow.GetString('TABLEAUVARIATION') = 'X', 2);
  AssigneImage( IMCommentaire, ATobRow.GetString('MILLESIME') = 'X', 3);
  AssigneImage( IMVisa, ATobRow.GetString('VISAREVISION') = 'X', 4);

  LPonderationCycle.Caption := TraduireMemoire('Pondération cycle') + ' : ';
  LEtatCycle.Caption := TraduireMemoire('<< Aucun >>');
  lStCycle := ATobRow.GetString('CYCLEREVISION');
  if lStCycle <> '' then
  begin
    lTobCycle := FTobListeCycle.FindFirst(['CCY_CODECYCLE'], [lStCycle], False);
    if lTobCycle <> nil then
    begin
      // Calcul de la pondération
      lDbTotalCompte := Arrondi(ATobRow.GetDouble('TOTDEBE') - ATobRow.GetDouble('TOTCREE'), 2);
      lDbTotalCycle  := Arrondi(lTobCycle.GetDouble('TOTAL'), 2);
      lDBTotalPonderation := 0;

      if (( lDbTotalCycle > 0 ) and ( lDbTotalCompte > 0 )) or
         (( lDbTotalCycle < 0 ) and ( lDbTotalCompte < 0 )) then
        lDBTotalPonderation := ( Abs(lDbTotalCompte) * 100 / (Abs( lDbTotalCycle )) );

      // Affichage de la pondération
      LPonderationcycle.Caption := LPonderationcycle.Caption + StrFMontant( lDbTotalPonderation , 15, V_PGI.OkdecV, '', True ) + ' %';

      // Affichage de l'état du cycle
      LEtatCycle.Caption := RechDom('CREVETATCYCLE', lTobCycle.GetString('CCY_ETATCYCLE'), False);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.AssigneImage (vImage : TImage; vBoAffiche : Boolean; vInID : integer);
begin
  if V_Pgi.ModeTSE then
  begin
    if vBoAffiche then
      vImage.Picture.Assign( IMOui.Picture.BitMap )
    else
      vImage.Picture.Assign( IMNon.Picture.BitMap );
  end
  else
  begin
    if not vBoAffiche then
      vImage.Picture.Assign( IMNon.Picture.BitMap )
    else
    begin
      case vInId of
      1 : vImage.Picture.Assign( IM1.Picture.BitMap );
      2 : vImage.Picture.Assign( IM2.Picture.BitMap );
      3 : vImage.Picture.Assign( IM3.Picture.BitMap );
      4 : vImage.Picture.Assign( IM4.Picture.BitMap );
      else
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/05/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnPostDrawCellFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : Lettrage Manuel d' un compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.LETTRAGEMClick(Sender: TObject);
{$IFDEF COMPTA}
var
  R               : RLETTR;
  lAction         : TActionFiche;
  lTobGene        : Tob;
  lStEtatLettrage : string;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTobGene := GetO(FListe) ;
  if lTobGene = nil then Exit;

  FillChar(R, Sizeof(R), #0);
  lAction         := taModif;
  R.General       := lTobGene.GetString('G_GENERAL');
  R.Auxiliaire    := '';
  R.Appel         := tlMenu;
  R.CritDev       := V_PGI.DevisePivot;
  R.DeviseMvt     := V_PGI.DevisePivot;
  R.GL            := nil;
  lStEtatLettrage := 'AL';
  R.CritMvt       :=' AND ( E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL" ) ' ;
  CEtudieModeLettrage(R) ;
  LettrageManuel(R, True, lAction) ;
  BCherche.Click;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : Lettrage Automatique
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.LettrageAClick(Sender: TObject);
var lTobGene : TOB ;
begin
  lTobGene := GetO(FListe) ;
  if lTobGene = nil then Exit;
{$IFDEF COMPTA}
  if not CEstAutoriseDelettrage(False, False) then Exit ;
  RapprochementAuto( lTobGene.GetString('G_GENERAL'), '' );
  BCherche.Click;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 28/07/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.DelettreTotal(Sender: TObject);
begin
 if not CEstAutoriseDelettrage(False, False) then Exit ;

 if PGIAsk('Attention, vous allez supprimer tout le lettrage du compte.' + #13 + #10 +
            'Confirmez-vous le traitement ?', 'Delettrage automatique') = mrYes then
 begin
   TRANSACTIONS(TransactionDeLettrageTotal,1) ;
   RefreshPclPGE;
 end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickPasseJustifSolde(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  OperationsSurComptes(ATobRow.GetString('G_GENERAL'), FExoDate.Code, '', '', False, '', True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.TransactionDeLettrageTotal ;
var lTobGene : Tob;
begin
  lTobGene := GetO(FListe);
  if lTobGene = nil then Exit;
{$IFDEF COMPTA}
  CExecDeLettrage( lTobGene.GetString('G_GENERAL'), '');
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 28/07/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.DelettreMvtExoRef(Sender: TObject);
begin
  if not CEstAutoriseDelettrage(False, False) then Exit ;

  if PgiAsk('Attention, vous allez supprimer tout le lettrage de l''exercice de référence.' + #13 + #10 +
            'Confirmez vous le traitement ?', 'Delettrage complet pour l''exercice de référence') = MrYes then
  begin
    TRANSACTIONS(TransactionDelettreMvtExoRef,1) ;
    RefreshPclPGE;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.TransactionDelettreMvtExoRef ;
var lTobGene : Tob;
begin
  lTobGene := GetO(FListe);
  if lTobGene = nil then Exit;
{$IFDEF COMPTA}
  CExecDelettreMvtExoRef( lTobGene.GetString('G_GENERAL'), '');
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGCD( Sender : TObject) ;
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CPLanceFiche_GCDOperation('' , 'GEN;CREANCE;;' + lTOB.GetString('G_GENERAL') + ';' + lTOB.GetString('SOLDEE') ) ;
  RefreshPclPge ;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPCONSGENE.OnClickTableauGCD( Sender : TObject) ;
begin
{$IFDEF COMPTA}
  CPLanceFiche_CPMULGCD() ;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Saisie d'une Piece
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.SaisieEcrClick(Sender: TObject);
{$IFDEF COMPTA}
  var OldSC: Boolean;
{$ENDIF}
begin
{$IFDEF COMPTA}
  OldSC := VH^.BouclerSaisieCreat;
  VH^.BouclerSaisieCreat := False;
  MultiCritereMvt(taCreat, 'N', False);
  VH^.BouclerSaisieCreat := OldSC;
  BCherche.Click;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Saisie d'un Bordereau
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.SaisieBorClick(Sender: TObject);
begin
{$IFDEF COMPTA}
  SaisieFolio(taModif);
  BCherche.Click;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.SaisieTresoClick(Sender: TObject);
{$IFDEF COMPTA}
{$IFNDEF CCMP}
var lTob : Tob;
{$ENDIF}
{$ENDIF}
begin
{$IFDEF COMPTA}
  {$IFNDEF CCMP}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CPLanceFiche_SaisieTresorerie(lTob.GetString('G_GENERAL'));
  {$ENDIF}
{$ENDIF}

(*
     Q := OpenSQL(' select distinct J_CONTREPARTIE '  +
                'from JOURNAL, BANQUECP '           +
                'where J_NATUREJAL ="BQE" '         +
                'and BQ_GENERAL=J_CONTREPARTIE '    +
                'and BQ_DEVISE="'                   + V_PGI.DevisePivot + '" ' +
                ' an
                d J_CONTREPARTIE is not null' ,
                true );
      *)

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Ouvre la fiche du compte en affichant l' onglet INFORMATIONS
suite ........ : pour consulter le BLOC NOTE
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickCommentCpte(Sender: TObject);
var lTob : Tob;
    lAction : TActionFiche;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  lAction := TaModif;
  FicheGene( nil, '', lTob.GetString('G_GENERAL'), lAction, 2);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Cumuls du compte général
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.CumulsGENEClick(Sender: TObject);
var lTob : TOB ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CumulCpteMensuel(fbGene, lTob.GetString('G_GENERAL'), lTob.GetString('G_LIBELLE'), VH^.Encours);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : Historique du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.HistoCpteClick(Sender: TObject);
var lTob : TOB ;
begin
 lTob := GetO(FListe) ;
 if lTob = nil then Exit;
 CC_LanceFicheHistoCpte( lTob.GetString('G_GENERAL') );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/09/2002
Modifié le ... :   /  /
Description .. : Mouvements analytiques du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.AnalytiquesClick(Sender: TObject);
{$IFDEF COMPTA}
var //lAction  : TActionFiche;
    lTob     : TOB ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritedt.CritEdt), #0);

  // GCO - 25/09/2007 - FQ 21085
  ACritEdt.CritEdt.Exo     := FExoDate;
  ACritEdt.CritEdt.Date1   := FExoDate.Deb;
  ACritEdt.CritEdt.Date2   := FExoDate.Fin;
  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;
  // FIN GCO - 25/09/2007

  ACritEdt.CritEdt.SCpt1   := lTob.GetString('G_GENERAL');
  //lAction := taModif;
  //if not FOkCreateModif then
  //  lAction := taConsult;
  TheData := ACritEdt;
  MultiCritereAnaZoom(TaModif, ACritEdt.CritEdt);
  TheData := nil
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/10/2002
Modifié le ... :   /  /
Description .. : Rapprochement Comptabilité / Immo
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickRapproComptaImmo(Sender: TObject);
begin
{$IFDEF AMORTISSEMENT}
    EtatRapprochementCompta;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickRapproComptaGCD(Sender: TObject);
begin
{$IFDEF COMPTA}
  CPLanceFiche_CPMULGCD('GCDCOMPTA') ;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... :   /  /
Description .. : Pointage du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.PointageGClick(Sender: TObject);
{$IFDEF COMPTA}
var
  lTOB : TOB ;
  Arg  : string;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  {JP 30/07/07 : Gestion de l'appel du nouveau pointage}
  if EstSpecif('51213') and (ctxPcl in V_PGI.PGIContexte) then Arg := ''
  else if MsgPointageSurTreso then Arg := CODENEWPOINTAGE + ';';

  if lTob.GetString('G_GENERAL') <> '' then
    CPLanceFiche_PointageMul( Arg + lTob.GetString('G_GENERAL') + ';' );

  BCherche.Click;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... :   /  /
Description .. : Pointage en cours
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.PointageEnCoursClick(Sender: TObject);
{$IFDEF COMPTA}
var lTOB : TOB ;
    lQuery : TQuery;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  lQuery := nil;
  try
    try
      {JP 12/07/07 : FQ 21045 : nouveau pointage}
      lQuery := OpenSql('SELECT EE_GENERAL, EE_REFPOINTAGE, EE_DATEPOINTAGE, EE_NUMERO, EE_ORIGINERELEVE ' +
                         'FROM EEXBQ WHERE EE_GENERAL = "' + lTob.GetString('G_GENERAL') + '"' +
                         'ORDER BY EE_DATEPOINTAGE DESC', True);

      if not lQuery.Eof then
      begin
        if EstSpecif('51213') and (ctxPcl in V_PGI.PGIContexte) then
          CPLanceFiche_Pointage( lQuery.FindField('EE_GENERAL').AsString + ';' +
                                 lQuery.FindField('EE_DATEPOINTAGE').AsString + ';' +
                                 lQuery.FindField('EE_REFPOINTAGE').AsString)
        else
          CPLanceFiche_PointageRappro('ACTION=MODIFICATION;' +
                                      lQuery.FindField('EE_GENERAL').AsString + ';' +
                                      lQuery.FindField('EE_DATEPOINTAGE').AsString  + ';' +
                                      lQuery.FindField('EE_REFPOINTAGE').AsString  + ';' +
                                      lQuery.FindField('EE_NUMERO').AsString + ';' +
                                      lQuery.FindField('EE_ORIGINERELEVE').AsString + ';');
      end
      else
        PgiInfo('Aucune référence de pointage n''a été trouvée.', 'Pointage en cours');

      BCherche.Click;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Attention');
    end;

  finally
    Ferme( lQuery );
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.ConsAuxiliaireClick(Sender: TObject);
var lTOB : TOB ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CPLanceFiche_CPCONSAUXI(lTob.GetString('G_GENERAL')) ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... :   /  /
Description .. : Liste des immobilisations du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.ListeImmoClick(Sender: TObject);
{$IFDEF AMORTISSEMENT}
var lTOB : TOB ;
{$ENDIF}
begin
{$IFDEF AMORTISSEMENT}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  AMLanceFiche_ListeDesImmobilisations ( lTob.GetString('G_GENERAL') , False, taModif );
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... :   /  /
Description .. : Détail du compte ICC
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickDetailIcc(Sender: TObject);
var lTOB : TOB ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  AGLLanceFiche('CP', 'ICCFICHEGENERAUX', '', lTob.GetString('G_GENERAL'), 'ACTION=MODIFICATION;' +
  DateToStr(FExoDate.Deb) + ';' +  DateToStr(FExoDate.Fin) );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickCalculIcc(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  AGLLanceFiche('CP', 'ICCPARAMETRE', '', '', DateToStr(FExoDate.Deb) + ';' +
                DateToStr(FExoDate.Fin) + ';' + lTob.GetString('G_GENERAL'));
end;

{$IFDEF COMPTA}
{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. : Rapprochement Compta / CRE
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickRappECRCRE(Sender: TObject);
begin
     EtatRapproCRECompta ; //Rapprochement Compta/CRE
end;
{$ENDIF}


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... :   /  /
Description .. : Calcul de reboursement d' emprunt
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickEmpruntCRE(Sender: TObject);
var lTOB : TOB ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  AGLLanceFiche('FP', 'FMULEMPRUNT', '', '', lTob.GetString('G_GENERAL'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... : 18/02/2004
Description .. : Justificatif de solde TID / TIC ou Tiers
Suite ........ : GCO - 1/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickJustifSolde(Sender: TObject);
{$IFDEF COMPTA}
var lTob : TOB ;
    lStNatureGene : string;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  lStNatureGene := lTob.GetString('NATUREGENE');

  if ( lStNatureGene = 'TIC') or
     ( lStNatureGene = 'TID') or
     ((lStNatureGene = 'DIV') and (lTob.GetString('LETTRABLE') = 'X')) then
  begin
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    PrepareArgumentGLG( True );
    ACritEdt.CritEdt.Cpt1 := lTob.GetString('G_GENERAL');
    ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;
    TheData := ACritEdt;
    CPLanceFiche_CPGLGENE('');
    TheData := nil;
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSGENE.PrepareArgumentBAL;
begin
  // Exercice
  ACritEdt.CritEdt.Exo.Code := FExoDate.Code;
  // Début de la Date Comptable
  ACritEdt.CritEdt.Date1 := FExoDate.Deb;
  // Fin de la Date Comptable
  ACritEdt.CritEdt.Date2 := FExoDate.Fin;
  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := 'N;';
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/09/2004
Modifié le ... : 25/05/2005
Description .. : Impression de la Balance Générale
Suite ........ : 25/05/2005 - FQ 15843 - Balance Générale avec tout les comptes
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBalGene(Sender: TObject);
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := G_General.Text;
  ACritEdt.CritEdt.Cpt2 := G_General_.Text;
  if (G_CycleRevision.Text <> '') then
    ACritEdt.CritEdt.BAL.CycleDeRevision := G_CycleRevision.Text;
    
  TheData := ACritEdt;
  CPLanceFiche_BalanceGeneral;
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2002
Modifié le ... : 20/08/2004
Description .. : Balance Générale par auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBalGeneParAuxi(Sender: TObject);
{$IFDEF COMPTA}
var lTob : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := lTob.GetString('G_GENERAL');
  ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;
  TheData := ACritEdt;
  CPLanceFiche_BalanceGenAuxi;
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/08/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBalAnaSurAxe1( Sender : TObject );
begin
  LancementBALGENEPARANA( fbAxe1 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/08/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBalAnaSurAxe2( Sender : TObject );
begin
  LancementBALGENEPARANA( fbAxe2 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/08/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBalAnaSurAxe3( Sender : TObject );
begin
  LancementBALGENEPARANA( fbAxe3 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/08/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBalAnaSurAxe4( Sender : TObject );
begin
  LancementBALGENEPARANA( fbAxe4 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/08/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBalAnaSurAxe5( Sender : TObject );
begin
  LancementBALGENEPARANA( fbAxe5 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/08/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.LancementBALGENEPARANA( vFichierBase : TFichierBase );
{$IFDEF COMPTA}
var lTob  : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1    := lTob.GetString('G_GENERAL');
  ACritEdt.CritEdt.Cpt2    := ACritEdt.CritEdt.Cpt1;
  ACritEdt.CritEdt.BAL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_BalanceGenAnal;
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2004
Modifié le ... : 19/02/2004
Description .. :
Suite ........ : GCO - 19/02/2004
Suite ........ : Utilisation du CritEDt pour le GL
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
Procedure TOF_CPCONSGENE.PrepareArgumentGLG ( vBoJustifSolde : Boolean );
begin
  // Exercice
  ACritEdt.CritEdt.Exo.Code := FExoDate.Code;
  // Début de la Date Comptable
  ACritEdt.CritEdt.Date1 := FExoDate.Deb;
  // Fin de la Date Comptable
  ACritEdt.CritEdt.Date2 := FExoDate.Fin;
  // Justificatif de Solde
  ACritEdt.CritEdt.GL.EnDateSituation := (vBoJustifSolde = True);

  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := 'N;';

  // Ecritures Valides
  ACritEdt.CritEdt.Valide := '';

  // Ecritures Lettrées
  ACritEdt.CritEdt.GL.Lettrable := 0;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2002
Modifié le ... : 18/02/2004
Description .. : Grand livre général par auxiliaire
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLGeneParAuxi(Sender: TObject);
{$IFDEF COMPTA}
var lTob : TOB ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1   := lTob.GetString('G_GENERAL');
  ACritEdt.CritEdt.Cpt2   := ACritEdt.CritEdt.Cpt1;

  TheData := ACritEdt;
  CPLanceFiche_CPGLGeneParAuxi('');
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... : 18/02/2004
Description .. : Grand livre du compte général
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLGene(Sender: TObject);
{$IFDEF COMPTA}
var lTob : TOB ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1 := lTob.GetString('G_GENERAL');
  ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;

  if (G_CycleRevision.Text <> '') then
    ACritEdt.CritEdt.GL.CycleDeRevision := G_CycleRevision.Text;

  TheData := ACritEdt;
  CPLanceFiche_CPGLGene('');
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLGeneParQte(sender: TObject);
{$IFDEF COMPTA}
var lTob : TOB ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1 := lTob.GetString('G_GENERAL');
  ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;

  TheData := ACritEdt;
  CPLanceFiche_CPGLGeneParQte('');
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... :   /  /
Description .. : Grand livre Analytique sur Axe 1..5
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLAnaSurAxe1(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe1 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLAnaSurAxe2(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe2 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLAnaSurAxe3(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe3 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLAnaSurAxe4(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe4 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickGLAnaSurAxe5(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe5 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/08/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.LancementGLGENEPARANA( vFichierBase : TFichierBase );
{$IFDEF COMPTA}
var lTob : Tob ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGLG( False );
  ACritEdt.CritEdt.Cpt1   := lTob.GetString('G_GENERAL');
  ACritEdt.CritEdt.Cpt2   := ACritEdt.CritEdt.Cpt1;
  ACritEdt.CritEdt.GL.Axe := fbToAxe( vFichierBase );

  TheData := ACritEdt;
  CPLanceFiche_CPGLGeneParAna;
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickNoteCtrlCompte(Sender: TObject);
begin
  TRic.CPLanceEtat_NoteCtrlCompte( FInfoGene.General );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... :   /  /
Description .. : Mise à jour du Caption de l' écran en fonction de la recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.MiseAJourCaptionEcran;
begin
  // GCO - 01/10/2007 - FQ 21508
  if FTobBDSE <> nil then
    Ecran.Caption := TraduireMemoire('Consultation des comptes généraux du ') + ' ' +
                     DateToStr(FTobBDSE.GetDateTime('BSI_DATE1')) +
                     TraduireMemoire('au') + ' ' +
                     DateToStr(FTobBDSE.GetDateTime('BSI_DATE2'))
  else
    Ecran.Caption := TraduireMemoire('Consultation des comptes généraux') + ' : ' +
                     RechDom('TTEXERCICE', FExoDate.Code, False);
  UpDateCaption(Ecran);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Analyse Complémentaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.ANALREVISClick(Sender: TObject);
var
 lWhereSQL : string;
 lTOB      : TOB;
begin
  inherited;
  lTob := GetO(FListe) ;
  if lTob = nil then Exit ;
  lWhereSQL := ' WHERE E_GENERAL LIKE "%' + lTob.GetString('G_GENERAL') + '%" AND E_QUALIFPIECE="N"' ;
  AGLLanceFiche('CP', 'CPCONSULTREVIS', '', '', lWhereSQL);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/08/2004
Modifié le ... : 25/11/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickEtatRapprochement(Sender: TObject);
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob <> nil then
    BEGIN
    If OkNewPointage Then CC_LanceFicheEtatRapproDet(lTob.GetString('G_GENERAL'))
                     Else CC_LanceFicheEtatRapproDetV7(lTob.GetString('G_GENERAL'));
    END ;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2004
Modifié le ... : 25/11/2004
Description .. : Justificatif des soldes bancaires
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickJustifSoldeBQE( Sender : TObject );
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob <> nil then
  begin
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    ACritEdt.CritEdt.Cpt1     := lTob.GetString('G_GENERAL');
    ACritEdt.CritEdt.Cpt2     := ACritEdt.CritEdt.Cpt1;
    ACritEdt.CritEdt.Exo.Code := FExoDate.Code; // Exercice
    ACritEdt.CritEdt.Date1    := FExoDate.Deb;  // Date Début
    ACritEdt.CritEdt.Date2    := FExoDate.Fin;  // Date Fin
    TheData := ACritEdt;
    CC_LanceFicheJustifPointage;
    TheData := nil;
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnElipsisClickG_General(Sender: TObject);
begin
  LookUpList( THEdit(Sender),
              'Comptes généraux',
              'GENERAUX',
              'G_GENERAL' ,
              'G_LIBELLE',
              CGenereSQLConfidentiel('G'),
              'G_GENERAL' ,
              True,
              1 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/12/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnExitG_General(Sender: TObject);
begin
  if (Trim(THEdit(Sender).Text) = '') or (TestJoker(THEdit(Sender).Text)) then Exit;

  if Length(THEdit(Sender).Text) < VH^.Cpta[fbGene].Lg then
    THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbGene);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickBEffaceRevision(Sender: TObject);
begin
  GTypePonderation.Value := 'SOLDE'; // Init à "SOLDE"
  ZCPonderation.ItemIndex    := -1;
  GValeurPonderation.Value   := 0;
  GBasePonderation.Text      := '';

  if GTypeVariation.Enabled then
    GTypeVariation.ItemIndex := -1;
  ZCVariation.ItemIndex      := -1;
  GValeurVariation.Value     := 0;
  GNatureVariation.Value     := 'VALEUR';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/01/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnElipsisClickGBasePonderation(Sender: TObject);
begin
  LookupList(GBasePonderation, TraduireMemoire('Recherche d''une rubrique'),'RUBRIQUE','RB_RUBRIQUE','RB_LIBELLE',
            'RB_CLASSERUB <> "CDR"', 'RB_RUBRIQUE', True, 7701);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CPCONSGENE.CompareOperateur(vValeur, vValeurAComparer: Double; vStOperateur: string): Boolean;
begin
  if vStOperateur = 'EGA' then
    Result := ( vValeur = vValeurAComparer)
  else
    if vStOperateur = 'IEG' then
      Result := ( vValeur <= vValeurAComparer)
    else
      if vStOperateur = 'INF' then
        Result := ( vValeur < vValeurAComparer)
      else
      if vStOperateur = 'SEG' then
        Result := ( vValeur >= vValeurAComparer)
      else
        if vStOperateur = 'SUP' then
          Result := ( vValeur > vValeurAComparer)
        else
          Result := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnChangeGTypePonderation(Sender: TObject);
begin
  // Activation de la base de ponderation sur recherche sur le solde uniquement
  GBasePonderation.Enabled := (GTypePonderation.Value = 'SOLDE');
  GBasePonderation.Color := IIF(GBasePonderation.Enabled, ClWindow, ClBtnFace);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnChangeGBasePonderation(Sender: TObject);
//var lValeur : Double;
begin
  //lValeur := 0;
  // GCO - 07/06/2007 - FQ 20613
  // if GValeurPonderation.Text <> '' then
  //   lValeur := Valeur(GValeurPonderation.Text);
//  lValeur := GValeurPonderation.Value;

  if GBasePonderation.Text <> '' then
  begin
    GValeurPonderation.NumericType := ntPercentage;
    //GCO - 07/06/2007 - FQ 20613
    //GValeurPonderation.Value := GValeurPonderation.Value + '%';
    //if GValeurPonderation.Text <> '' then
      //GValeurPonderation.Text := StrFMontant(lValeur, 15, V_PGI.OkDecV, '', True) + '%';
  end
  else
  begin
    GValeurPonderation.NumericType := ntGeneral;
    //GCO - 07/06/2007 - FQ 20613
    //GValeurPonderation.Value := GValeurPonderation.Value  + '';
    //if GValeurPonderation.Text <> '' then
      //GValeurPonderation.Text := StrFMontant(lValeur, 15, V_PGI.OkDecV, '', True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnChangeGNatureVariation(Sender: TObject);
begin
  // GCO - 05/11/2007 - FQ 21776 + suppresion de la FQ 21578 (Ajout du *100)
  // pour que ça fonctionne
  (*
  if GNatureVariation.Value = 'POURCENTAGE' then
  begin
    GValeurVariation.NumericType := ntPercentage;
    //GCO - 07/06/2007 - FQ 20613
    //GValeurVariation.Text := StrFMontant(Valeur(GValeurVariation.Text)/100, 15, V_PGI.OkDecV, '', True) + '%';
    //GValeurVariation.Value := GValeurVariation.Value + '%';
  end
  else
  begin
    GValeurVariation.NumericType := ntGeneral;
  end;*)
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/02/2005
Modifié le ... : 07/02/2005    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.TripotageGSituationCompte;
begin
  GSituationCompte.Items[1] := 'Mouvementés sur l''exercice N mais sans mouvements sur N-1';
  GSituationCompte.Items[2] := 'Sans mouvement sur l''exercice N mais mouvementés sur N-1';
  GSituationCompte.Items[3] := 'Non soldés sur l''exercice N mais soldés sur N-1';
  GSituationCompte.Items[4] := 'Soldés sur l''exercice N mais non soldés sur N-1';
  GSituationCompte.Items[5] := 'N''ayant pas le même sens sur l''exercice N que sur N-1';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2006
Modifié le ... :   /  /
Description .. : GCO - FQ 17496
Mots clefs ... :
*****************************************************************}
(*
procedure TOF_CPCONSGENE.OnChangeG_CycleRevision(Sender: TObject);
begin
  TMPGCycleRevision.Text := 'X' + G_CycleRevision.Text;
end;
*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/05/2006
Modifié le ... :   /  /    
Description .. : FQ 17496 - impossible de faire autrement
Mots clefs ... :
*****************************************************************}
(*
procedure TOF_CPCONSGENE.InitSelectFiltre(T: TOB); MBAMF
begin
  inherited;
  G_CycleRevision.Value := Copy(TMPGCycleRevision.Text, 2, Length(TMPGCycleRevision.Text)-1);
end; *)

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPCONSGENE.OnClickImprimer(Sender: TObject);
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.AvecComparatif := True; // Balance Comparative
  if FExoDate.Code = VH^.Encours.Code then
   ACritEdt.CritEdt.ExoComparatif := VH^.Precedent
  else ACritEdt.CritEdt.ExoComparatif := VH^.EnCours;

  ACritEdt.CritEdt.Cpt1 := G_General.Text;
  ACritEdt.CritEdt.Cpt2 := G_General_.Text;
  if (G_CycleRevision.Text <> '') then
    ACritEdt.CritEdt.BAL.CycleDeRevision := G_CycleRevision.Text;

  TheData := ACritEdt;
  CPLanceFiche_BalanceGeneral ('BAR');
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickFeuilleExcel(Sender: TObject);
begin
  CPLanceFiche_CPIMPFICHEXCEL ('FTS;' + GetControlText ('G_GENERAL') + ';' + GetControltext ('G_CYCLEREVISION'));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 10/11/2006
Modifié le ... : 05/01/2007
Description .. : Appel note de travail
Suite ........ : Arguments : Compte général
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickNoteTravail( Sender: TObject );
{$IFDEF COMPTA}
var TobC : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  Tobc := GetO(FListe);
  if Tobc = nil then Exit;
  CPLanceFiche_CPNoteTravail( TobC.GetString('G_GENERAL'));
  BCherche.Click;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/11/2006
Modifié le ... : 05/01/2007
Description .. : Appel du tableau des variations
Suite ........ : Argument : Compte général
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickTableauVariation( Sender : TObject );
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPTableauVar( lTob.GetString('G_GENERAL'));
  BCherche.Click;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickRechercheEcritures(Sender: Tobject);
begin
  if CtxPcl in V_Pgi.PgiContexte then
    CPLanceFiche_CPRechercheEcr( False )
  else
    MultiCritereMvt(taConsult,'N', False );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CPCONSGENE.OnClickDRF ( Sender : TObject );
begin
  LanceFiche_DRF( OkLanceDRF );
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickCommentaire( Sender: TObject );
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPREVDocTravaux(lTob.GetString('G_GENERAL'), '', VH^.Encours.Code, 3);
  BCherche.Click;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2007
Modifié le ... : 07/02/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.OnClickEtatCycle(Sender: Tobject);
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPREVDocTravaux(lTob.GetString('G_GENERAL'), '', VH^.Encours.Code);
  BCherche.Click;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
// GP le 15/07/2008 EVOL AGL 7.2
(*
procedure TOF_CPCONSGENE.OnMouseEnterHyperLinkLabel(Sender: TObject; Enter : Boolean);
begin
  if Enter then
   THLabel(Sender).Font.Color := ClBlue
 else
   THLabel(Sender).Font.Color := ClBlack;
end;
*)


procedure TOF_CPCONSGENE.OnMouseEnterHyperLinkLabel(Sender: TObject);
begin
  THLabel(Sender).Font.Color := ClBlue
end;

procedure TOF_CPCONSGENE.OnMouseLeaveHyperLinkLabel(Sender: TObject);
begin
  THLabel(Sender).Font.Color := ClBlack;
end;


procedure TOF_CPCONSGENE.OnClickCutOff(Sender: TObject);
{$IFDEF COMPTA}
var lStGen : string ;
    lTOB   : TOB ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob   := GetO(FListe) ;
  lStGen := Copy(lTob.GetString('G_GENERAL'), 1, 3);
  if ( lStGen = '486' ) or ( lStGen = '487' ) then
    CPLanceFiche_CPMULCUTOFF()
  else
    CPLanceFiche_CPMULCUTOFF(lTOB.GetString('G_GENERAL'));
  // GCO - 05/09/2006 - FQ 18701
  RefreshPclPge;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CPCONSGENE.OnClickControleTva(Sender: TObject);
var lStGene : string;
begin
  // GCO - 26/07/2007 - FQ 20184
  lStGene := Copy(FInfoGene.General, 1, 5);

  if (lStGene = '44566') or (lStGene = '44571') then
  begin // Tva Déductible ou Tva collextée
    CPLanceFiche_CPCONTROLTVA( IIF(lStGene = '44566', 'DED;', 'COLL;') );
  end
  else
    CPLanceFiche_CPCONTROLTVA(''); // Tva Déductible
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////

{$IFDEF COMPTA}
{$IFNDEF CCMP}
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 29/03/2007
Modifié le ... :   /  /
Description .. : Répartition analytique
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSGENE.RepartitionAnaClick(Sender: TObject);
var lTob : TOB ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.Cpt1    := lTob.GetString('G_GENERAL');
  ACritEdt.CritEdt.Cpt2    := ACritEdt.CritEdt.Cpt1;
  ACritEdt.CritEdt.Date1   := VH^.Encours.Deb;
  ACritEdt.CritEdt.Date2   := VH^.Encours.Fin;
  ACritEdt.CritEdt.DateDeb := VH^.Encours.Deb ;
  ACritEdt.CritEdt.DateFin := VH^.Encours.Fin ;
  ACritEdt.CritEdt.Exo     := VH^.Encours;
  ACritEdt.CritEdt.SCpt1   := ACritEdt.CritEdt.Cpt1;

  TheData := ACritEdt;
  CPLanceFiche_RepartitionAnalytique(taModif, lTob.GetString('G_GENERAL'));
  TheData := nil;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSGENE.InitBalanceSituation;
var lStBDSE : string;
    lStBDSP : string;
begin
    lStBDSE := GetParamSocSecur('SO_CPBDSENCOURS', '', True);
    if lStBDSE <> '' then
    begin
      if not ChargeBalanceSituation( lStBDSE, True ) then
        PgiError('Balance de situation ' + lStBDSE + ' non trouvée', 'Erreur dans les paramètres société');
    end;

    lStBDSP := GetParamSocSecur('SO_CPBDSPRECEDENT', '', True);
    if lStBDSP <> '' then
    begin
      if not ChargeBalanceSituation( lStBDSP, False ) then
        PgiError('Balance de situation ' + lStBDSP + ' non trouvée', 'Erreur dans les paramètres société');
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CPCONSGENE.ChargeBalanceSituation  ( vStCodeBal : string; vBoSurN : Boolean ) : Boolean;
var lQuery : TQuery;
begin
  Result := False;
  lQuery := OpenSql('SELECT BSI_DATE1, BSI_DATE2 FROM CBALSIT WHERE ' +
                    'BSI_CODEBAL = "' + vStCodeBal + '" ORDER BY BSI_CODEBAL', True);
  if not lQuery.Eof then
  begin
    Result := True;

    if vBoSurN then
    begin
      FTobBDSE := Tob.Create( vStCodeBAL, nil, -1);
      FTobBDSE.AddChampSupValeur('BSI_CODEBAL', vStCodeBal);
      FTobBDSE.AddChampSupValeur('BSI_DATE1', lQuery.FindField('BSI_DATE1').AsDateTime);
      FTobBDSE.AddChampSupValeur('BSI_DATE2', lQuery.FindField('BSI_DATE2').AsDateTime);
    end
    else
    begin
      FTobBDSP := Tob.Create( vStCodeBAL, nil, -1);
      FTobBDSP.AddChampSupValeur('BSI_CODEBAL', vStCodeBal);
      FTobBDSP.AddChampSupValeur('BSI_DATE1', lQuery.FindField('BSI_DATE1').AsDateTime);
      FTobBDSP.AddChampSupValeur('BSI_DATE2', lQuery.FindField('BSI_DATE2').AsDateTime);
    end;  
  end;
  Ferme(lQuery);
end;

(*
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CPCONSGENE.ChargeBalanceSituationPrecedent: Boolean;
var lQuery : TQuery;
begin
  Result := False;
  lQuery := OpenSql('SELECT BSI_DATE1, BSI_DATE2 FROM CBALSIT WHERE ' +
                    'BSI_CODEBAL = "' + FStBDSP + '" ORDER BY BSI_CODEBAL', True);
  if not lQuery.Eof then
  begin
    Result := True;
    FTobBDSE := Tob.Create('BDSE', nil, -1);
    FTobDBSE.AddChampSupValeur('BSI_CODEBAL', FStBDSP);
    FTobDBSE.AddChampSupValeur('BSI_DATE1', lQuery.FindField('BSI_DATE1').AsDateTime);
    FTobDBSE.AddChampSupValeur('BSI_DATE2', lQuery.FindField('BSI_DATE2').AsDateTime);
  end;
  Ferme(lQuery);

  FTobBDSP := Tob.Create('BDSP', nil, -1);
end; *)

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPCONSGENE ] ) ;
end.











        


