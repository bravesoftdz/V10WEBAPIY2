{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
unit MDispS5 ;

interface

uses
    TomIdentificationBancaire,
{$IFDEF TAXES}
    TomYMODELETAXE, // SDA le 12/02/2007
    YYMODELECATREG_TOF, //SDA le 28/02/2007
    YYMODELECATTYP_TOF, //SDA le 28/02/2007
    YYDETAILTAXES_TOF,//SDA le 28/02/2007
    YYCOMPTESPARCAT_TOF, //SDA le 02/03/2007
    YYTYPESTAXES_TOF, //SDA le 05/03/2007
{$ENDIF}
    YNewMessage,  // pour ShowNewMessage
{$IFDEF ADRESSE}
    PARAMADRE_TOM,    // Gestion des adresses.
    SAISIEADR_TOF,
{$ENDIF}
    EdtEtat,
{$IFDEF EAGLCLIENT}
    utileagl, // ConnecteHalley
    MenuOLX,
    MaineAGL,
    eTablette,
    About,
    Buttons,
    ComCtrls,
    Graphics,
    Hout,
    IniFiles,
    M3FP,
    Prefs,
    ulibzoom,
    {$IFDEF CONTREPARTIE}
    UtofReparationfic,  // ajout me
    {$ENDIF}
{$ELSE}
    MenuOLG,
    uedtcomp,
    Tablette,
    FE_Main,
    majTable, // ConnecteHalley
{$ENDIF EAGLCLIENT}
    BobGestion,
    CPGENEECRTP_TOF,
    Forms,
    Vierge,
    UTob,
    sysutils,
    dialogs,
    HMsgBox,
    HTB97,
    Classes,
    Controls,
    HPanel,
    Hctrls,
    UIUtil,
    ImgList,
    AglInitPlus,
    AGLInit,
    uTilGed,
    uMultiDossier,
    CPANOECR,
    HEnt1,        // TActionFiche
    GalOutil,
    licutil,      // CryptageSt, DayPass
    AnnOutils,    // JaiLeDroitBureau
    AglMail       // pour AglMailForm
    {$IFDEF NETEXPERT}
    ,UNetListe
    ,UAssistComsx
    ,uNEActions
    ,Echg_Code
    ,UAssistImport
    ,CPJALIMPORT_TOF
    {$ENDIF}
    ,reseau
    ,CPGestionAna_Tof //CPLanceGestionAna //SG6 08/12/2004
    ,CPTPayeurFacture_Tof
//    ,CPBalAgeeTPayeur_tof Lek 190106
    ,CPTPayeurGLAge_tof
    ,CPGLAGEVEN_tof  {Tiers Edition GL Lek 190106}
    ,CPEDTLegal_tof  {Edition Legal    Lek 190106}
    ,CPJustSold_Qr1_tof
    ,CpBalAgeeTp_Qr1_Tof
    ,CPBALOUVCLO_TOF
    ,CPBROUILLARDANA_TOF
    ,CPBROUILLARDBUD_TOF
    ,CPBROUILLARDECR_TOF
    ,CPJALCPTGENE_TOF
    ,CPJALOUVCLO_TOF
    ,AMSREGRO_TOF
    ,CPMODRESTANA_TOF
    ,CPRESTANABUD_TOF
    ,CPRESTANAECR_TOF
    ,CpCSectAna_tof
    ,CpCEcrAna_tof
    ,CPSAISIEPIECE_TOF
    ,CPMULMAJAFFAIRE_TOF
    ,MvtSect
    ,CPETATMARGES_TOF
    ,CpMultiBud_Qr1_tof      // Lek 270606 Multi budget
    ,CPGESTENG_TOF  // thl:en commentaire si cwas
    ,CPRAPAUTOENG_TOF
    ,CPSUIREAENG_TOF
    , dpTOFCETEBAC
    , ULibPointage
    , CPMULCETEBAC_TOF
    , CPMULEEXBQLIG_TOF
    , FPRAPPROCRE_TOF
    , CPRevMulParamCycl_TOF // CPLanceFiche_CPREVMULPARAMCYCL('');
    , CPPARAMEXTTVA_TOM
    , CPZOOM
//    , CPEDTSYNTH_TOF
    ,CPAVERTNEWRAPPRO_TOF
    ;

type
  TFMDispS5 = class(TDatamodule)
    ImageList: TImageList;
    IMREVISION: TImageList;

    procedure FMDispS5Create           ( Sender : TObject );
    procedure AfterChangeGroup         ( Sender : TObject );
    procedure ExercClick               ( Sender : TObject );
    procedure OnClickBStatutLiasse     ( Sender : TObject );
    procedure OnClickBStatutRevision   ( Sender : TObject );

    private
      BStatutLiasse  : TToolBarButton97;
      BStatutRevision : TToolBarButton97;

      NumModule : integer;
      procedure InitBStatutLiasse;
      procedure InitBStatutRevision;
    end;

Var FMDispS5 : TFMDispS5 ;
    MenuPCL : boolean ;
    vListeErreur : TStringList ;
    Init : boolean ;

{$IFDEF SCANGED}
procedure MyAfterImport (Sender: TObject; FileGuid: string; var Cancel: Boolean) ;
{$ENDIF}

function MyAglMailForm ( var Sujet : HString; var A, CC: String ; Body: HTStringList ; var Files: String ;
    ZoneTechDataVisible: boolean=True) : TResultMailForm ;

function  OkRevis : boolean ;
procedure ControleLettrageEnSaisie ;
function  AffichageEcranAccueil : boolean;
function LanceEcranAccueil( vstAccueil : string ) : Boolean;

Function  GetStRevis : String ;
function  OkDispatchPCL10 : Boolean;
procedure LanceEtatLibre ( NatureEtat : String ) ;
Procedure LanceListeLibre ( NatureEtat : String ) ;
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
procedure ChangeMenuDeca ;
Procedure AfterProtec ( sAcces : String ) ;
procedure AssignLesZoom ( LesImmos : boolean ) ;
procedure AfterChangeModule ( NumModule : integer ) ;

function  ChargeFavoris : boolean ;

procedure RemoveFromMenu ( const iTag : integer; const bGroup : boolean; var stAExclure : string );
procedure TripotageMenuModePCL ( const TNumModule : array of integer; var stAExclure : string );
procedure TripotageMenuAmortissement ( var stAExclure : string );
function IsOracle : Boolean;
procedure TripotageMenuModePGE ( const TNumModule : array of integer; var stAExclure : string );
procedure TripotageMenuInternat ( const TNumModule : array of integer; var stAExclure : string );
procedure TripotageMenu (NumModule : integer);
procedure AssignZoom ;
procedure InitApplication ;
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
Procedure InitLaVariablePGI;
procedure InitSerialisation;
procedure LanceMenuAuto;
function  ModeRecalcul : boolean ;
Procedure EnvoiBL;
Procedure ControlEnvoiBL;
procedure OnAfterUpdateBarreOutils ( Sender : Tobject );


//************************************************************************************************//
//************************************************************************************************//
//************************************************************************************************//
implementation

uses
  {$IFDEF MODENT1}
    CPTypeCons,
    CPVersion,
    CPProcGen,
  {$ENDIF MODENT1}
    Windows,
    Messages,
    Menus,
    Audit,
    wCommuns,
    CPMULTYPEVISA_TOF,
    CPMULCIRCUITBAP_TOF,
    CPMULAFFECTBAP_TOF,
    CPMAILBAP_TOM,
    CPMULBAP_TOF,
    CPHISTORIQUEMAIL_TOF,
    CPTYPEVISA_TOM,
    CPCIRCUITBAP_TOF,
    Commun, {Pour l'initialisation des agences, menu 3217}
    UTomBanque, {Fiche banque}
    TomAgence, {Agences bancaires}
    TRMULCOMPTEBQ_TOF, {Liste des comptes bancaires}
    CPMulRub_TOF,
    Rubrique_TOM,
    CPQRTRUBRIQUE_TOF,
    CPSUPPRUB_TOF,
    CPCONTROLERUB_TOF,
    CPTOTRUB_TOF,
    ESPCPQRIRPF115_TOF,  // TVA Espagnole
    UTOFMULRELFAC,     {JP 28/10/05 : Suppression de la fiche 2/3 RelevTie.Pas}
    CPMODIFECHEMP_TOF, {JP 28/10/05 : Suppression de la fiche 2/3 ModEcheMp.Pas}
    PLANGENE_TOF,      {JP 28/10/05 : Suppression de la fiche 2/3 QRPLANGE.Pas}
    PLANAUXI_TOF,      {JP 28/10/05 : Suppression de la fiche 2/3 QRPlanJo.Pas}
    PLANSECT_TOF,      {JP 28/10/05 : Suppression de la fiche 2/3 QRSectio.Pas}
    PLANJOUR_TOF,      {JP 28/10/05 : Suppression de la fiche 2/3 QRPlanJo.Pas}
    CPCUMULPERIODIQUE_TOF, {JP 24/06/05 : migration CWAS} { Ancien QRCumulG } // SBO 16/11/2005 : Suppression de la fiche 2 tiers
    CPBALAGEE_TOF,     // Nouvelle balance agée / ventilée { Ancien TofBalAgee }
    REPARTBU_TOM, {JP 21/11/05 : FQ 16521 : Remplacement de la fiche Delphi par la fiche Agl}
    CPRUBTLIB_TOF, //YMO 02/12/2005  appel par CplanceFiche
    CPVALIDEPIECE_TOF, // SBO Nouveau module de validation des pièces comptables
    CLIENSSOC_TOM,     // liaisons inter-sociétés
    uTOFCPMulGuide,
    CPMASQUEMUL_TOF, // SBO 14/11/2006 : mul des masques de saisie
    {--Lek 230106 sortir de EAGLCLIENT}
    BALBUDTEGEN_TOF,
    TOFEMPIMP,
    MulMvtBu_TOF, eSaisBud, SupEcrBU_TOF, MULBUDG_TOF, BUDGENE_TOM, MULBUDS_TOF, BUDSECT_TOM,
    MULBUDJ_TOF, SUPPRBUDG_TOF, SUPPRBUDS_TOF, SUPPRBUDJ_TOF, CroisCpt, TTCATJALBUD_TOT, CreRubBu, CtrBud,
    GenToBud, CodBuRup, CreCodBu, CreSecBu, Csebucat, ////
    MULVALBU_TOF, COPIEBUD_TOF, RevBuRea, COPIEBUDMU_TOF, OUVREFERME_TOF
    ,BalBudteSec_tof
    ,BalBudteGeS_tof
    ,BalBudteSeG_tof
    ,CpBalEcBudG_QR1_Tof
    ,CpBalEcBudS_QR1_Tof
    ,CpBalEcBudGS_QR1_Tof
    ,CpBalEcBudSG_QR1_Tof,
    SUIVCPTA_TOM,
    CPTVAMODIF_TOF,
    CPTVAACC_TOF,
    CPSUIVITVA_TOF,
{$IFDEF EAGLCLIENT}
    UTOFEnregStandard,
    CEGIDIEU, Hcompte,
    MULLETTR,
    AnulCloSERVEUR,
     CPVALTVA_TOF,
     PAYS_TOM, CPREGION_TOF, CPCODEPOSTAL_TOF, CPTVATPF_TOF, CPSTRUCTURE_TOF,
    CPCODEAFB_TOF, CPGENERAUX_TOM, CPTIERS_TOM, CPSECTION_TOM, CPJOURNAL_TOM,
    uTofConsEcr, CPMULANA_TOF,
    DEVISE_TOM,
    REFAUTO_TOM,
    REPARCLE_TOM, {JP 08/06/05 : migration CWAS}
    CPVIDESECTIONS_TOF, {JP 10/06/05 : migration CWAS}
    CPMODIFTABLELIBRE_TOF, {JP 10/06/05 : migration CWAS}
{$ELSE}
    VideSect,
    Refauto_Tom,  // ParamLib  BVE 25.06.07 FQ 20165
    EdtREtat,    // LanceEtat
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    PGIExec,
    ConsEcr,
    MULLETTR, AnulClo,
    CPTVATPF_TOF,
    CONTABON_TOM,
    TvaValid,
    Pays, Region, CodePost,
    DEVISE_TOM,
    CPGeneraux_TOM,
    CPTiers_TOM,
    CPJournal_TOM,
    CPSection_TOM,
    { FQ 20198 Structur} CPSTRUCTURE_TOF, { FQ 19684 Scenario,} CODEAFB,
    Extourne, Constant,
    Traduc, EdtDoc,
    PlanRef, RUPTURE, REPARCLE, EtapeReg, SuivDec,
    EtbMce, CreerSoc, InitSoc,  AssistSO, CopySoc,
    CtrlFic, CtrlDiv, ReIndex, TofRubImport,
    //calceuro,
    UTOFEnregStandard,
    ModTlEcr,
    MulAbo, MulAna,
    Edtlegal,
    EncaDeca, EcheModf,
    SuivTres, TVAEdite,  MulBds,
    CPAFAUXCPTA_TOF,
    PssToTli, // pas encore compatible eAGL
    CopiTabL, // pas encore compatible eAGL
    MulLog1,        {YMO 27/06/2006 FQ17895 Remise en place menu 'suivi d'activité' en CWAS ** 08/12/2006 Suppression menu}
  {$IFDEF CERIC}
    FonctionCERIC,            // Menu Spécifique CERIC
  {$ENDIF CERIC}
{$ENDIF EAGLCLIENT}
   VisuEtebac, {JP 01/06/07 : Branchement en eAGL}
   UlibEcriture ,
   {$IFDEF SCANGED}
    YNewDocument,
    UGedFiles,
  {$ENDIF SCANGED}
// FICHIERS COMMUNS
    CPQRTVAFAC_TOF,  //YMO 13/07/2006 DEV 3728 Suppression QrS1 sur états préparatoires de TVA
    CoutTier,        //YMO 03/07/2006 FQ 17850 Menu Estimation des coûts financiers dispo en CWAS
    CPQRTVACTRL_TOF, // YMO 25/01/2005 Branchement Etat de contrôle en 2 tiers
    eSaisieTr, {JP 13/01/06 : Suppression de SaisieTr}
    Ent1,
    uLibCalcEdtCompta,
    Choix,       // Choisir
    UTOFMULPARAMGEN, UtilSoc, MULDLETT, RELANCE_TOM,
    RappAuto, BonPayer, UTOFCPJALECR,
    UTOFCPGLGENE,
    UTOFCPGLAUXI,
    UTOFCPGLGESTION,
    CPGRANDLIVRETL_TOF, CPBALGEN_TOF, CPBALAUXI_TOF,
    CPBALANCETL_TOF, CRITSYNTH, CRITEDT, TofEcheancier, uTOFPointageMul, CPETATPOINTAGE_TOF,
    CPBALANAL_TOF, // Balance analytique
    CPBALGENANAL_TOF, CPBALGENAUXI_TOF, CPBALANALGEN_TOF, CPBALAUXIGEN_TOF, // Balances combinées
    uTofCPGLAna,
    Cloture, // cloture bi-compatible
    UTOFMODIFENTPIE,
    UTOFCPMULMVT,
    LGCOMPTE_TOF,
    CHGNATUREGENE_TOF,
    CPCONTRTVA_TOF,
    CPRAPPRODET_TOF,
    CORRESP_TOM,
    CPJUSTBQ_TOF, Valperio, ValJal, OuvreExo, CloPerio, ClotANA, AnulCAN, TofExpEtafi, TOFRECEP, TOFRRECU, TofRPoin,
    UTOFICCGENERAUX, UTOMICCGENERAUX, UTOFICCENATIONAUX, CPPRORATA_TOM, Prorata_tof, CPECRPRORATISEES_TOF, FichComm,
    CPTABLIBRELIB_TOF, CPCHOIXTALI_TOF, Synchro, EXERCICE_TOM, SOUCHE_TOM, USERGRP_TOM, Confidentialite_TOF, UTOMUTILISAT,
    utilPGI, Resulexo, CPTEUTIL, EntPGI, HalOLE, TomProfilUser, CPDPOINTAGEMUL_TOF, 
    uTofStdAlignePlan,
    ModiTali, // VL 26/01/2005 ModifSerieTableLibre
    BANQUECP_TOM,  // SBO 28/02/2005 Banquecp plus utilisé
    uTofConsAuxi, // CPLanceFiche_CPCONSAUXI
    uTofRechercheEcr, // CPLanceFiche_CPRechercheEcr
    ZJournalCentralisateur, // CPLanceFiche_JournalCentralisateur
    LibChpLi,     // FQ 16256 SBO 01/08/2005
    CPMULSUPPRANA_TOF, // SBO 26/08/2005 : remplacement fiche 2 tiers par fiche cwas // Axe
    CPAXE_TOM,         // SBO 26/08/2005 : remplacement fiche 2 tiers par fiche cwas // supprana
    Saisie,
    UTOFMULRELCPT {YMO 24/05/2006 Suppression Relcompt en 2T}
{$IFDEF AMORTISSEMENT}
    , AMLISTE_TOF
    , AMEDITION_TOF
    , AMANALYSE_TOF
    , AMINTEG_TOF
    , IMMOCPTE_TOM
    , AMCLOTURE_TOF
    , AMSuiviDPI_TOF
    , Outils
    , ImRapCpt
    , ImoClo
    , ImPlan
    , ImSaiCoef
//    , RepImmo
//    , CtrlImmo
    , AmVerif
    , ImOutGen
    , IntegEcr
    , ImOLE
    , ImEdCalc
    , ImpCegid
    , AmExport
    , AmExport_tof
    , AMLISTEARD_TOF
    , AMLISTEREGFR_TOF
{$ENDIF}
    , SaisBor
    , uTofSaisieEtebac
    , uTofSaisieVrac
    , uTofBalSitMul
    , ZReclassement       // Reclassement des écritures
    , ZCentral            // Journal centralisateur
    , uTofConsGene        // Consultation des comptes
    , uTofCPPurgeExercice // CPLanceFiche_CPPurgeExercice
    , SaisBal             // Saisie de balances N
    , AssistPL            // Assistant de comptabilité PGI Expert
    , GrpFavoris          // Favoris
    , Raz                 // Remise à zéro société
    , uTOFEtatsChaines    // CPLanceFiche_EtatsChaines
    , uObjEtatsChaines
    , GenerAbo            // Abonnements
    , CPReconAbo_TOF      // Reconduction des abonnements
  {$IFDEF EAGLCLIENT}
    , CPMulABo_TOF
    , Contabon_TOM
  {$ENDIF}
    , CPTRAITEMENTIAS_TOF // Traitement des AS14 (reventilation analytique)
    , CPRELANCECLIENT_TOF  // Relances clients
    , CalcScor             // calcul scoring
    , CroiseAxeTof         // Cube typé balance agée / ias14
    , ParamSoc
    , PwdDossier
    , LinkPublifi          // Lien avec Publifi
    , ReventilAna
    , UTOFEXTOURNE
    , CPREIMPUT_TOF
    , UTOTVENTILTYPE
    , UTofCutOff
    , LetRegul
    , NATCPTE_TOM
    {$IFNDEF EAGLCLIENT}
      ,YYBUNDLE_TOF
    {$ENDIF EAGLCLIENT}
    ,CLIENSETAB_TOM
    , UtofClotureDef
    , UTraitelibelle //ajout me 10/11/2005
    , SoldeCPT       // sbo 24/04/2006
    , CPCONTROLTVA_TOF
    , SaisUtil

{$IFDEF CISXPGI}
   ,URELEVINTER_TOF
{$ENDIF}

    // Révision
    , uLibRevision
    , CPStatutRevision_TOF
    , CREVInfoDossier_TOM
    , CPRevProgTravail_TOF

    // Suivi des règlements
    , CPMULENCADECA_TOF
    , UTOFCPSUIVIMP
    , UTofCPGenereMP
    , eSaisEff
    , CpteSav
    , CPMULPARAMGENER_TOF
    , CPMULEXPORT_TOF
    , CPGENECOMPEN_TOF
    , CPGENERENCADECA_TOF
    {$IFDEF EAGLCLIENT}
    , CPSUIVIACC_TOF
    , CPCFONBMP_TOF
    {$ELSE}
    , CFONBMP
    , AffBqe
    , MulSuivACC
    , MulSuivBAP
  {$ENDIF}
    , CPVALIDEECR_TOF
    , CPCONTROLELIASSE_TOF
    , CPTOTLIASSE_TOF
    , CPSTATUTDOSSIER_TOF
    , CPCONTROLCAISSE_TOF
    , CPLOIVENTILMUL_TOF
    , uImportBOB_TOF
    , CTableauBord
    , PLANBUDGET_TOF
    , PLANBUDSEC_TOF
    , PLANBUDJAL_TOF
    , CPGestionCreance
    , CPCFONB_TOF     // remplace CFONBAT
    , CPREIMPUTANA_TOF
    // DADS2
    ,UTOFPG_MULDADS2_TIERS    //LanceMul_DADS2Tiers
    ,UTOFPG_MULDADS2_HONOR    //LanceMul_DADS2Honoraires
    ,UTOFPG_DADS2_EDIT        //LanceFiche_DADS2Edition
    ,UTofPG_DADS2_PREP        //LanceFiche_DADS2Prep
    ,UTofPG_MulEnvoiSocial    //LanceMul_DADS2Envoi
    ,Revision                 //LanceRevision
    ,LETTREF                  //LettrageParCode (Lettrage par code de regroupement)
    ,uTOFPointageMulV7        //CPLanceFiche_PointageMulV7 (pointage manuel v2007)
    ,CPRAPPRODETV7_TOF        //CC_LanceFicheEtatRapproDetV7 (état rappro v2007
    {$IFDEF JOHN}
    , CPINTEGRECFONB_TOF
    {$ENDIF JOHN}
{$IFDEF CERTIFNF}
    ,CPJNALEVENT_TOF          //CPLanceFicheCPJNALEVENT Mul du journal des évenements
{$ENDIF}
    ,TRMULCIB_TOF             //TRLanceFiche_MulCIB
    ,REGLEACCRO_TOM           //TRLanceFiche_ParamRegleAccro
    ,TomCIB                   //TRLanceFiche_CIB
    ,Constantes               // tc_cib 
    ;

{$R *.DFM}

//************************************************************************************************//
//************************************************************************************************//
//************************************************************************************************//
{$IFDEF SCANGED}
procedure MyAfterImport (Sender: TObject; FileGuid: string; var Cancel: Boolean) ;
var
 ParGed : TParamGedDoc;
 NomFichier : string;
begin

  if FileGuid = '' then exit;

  // ---- Insertion dans la GED sans interaction utilisateur ----
  if TForm(Sender) is TFSaisieFolio then
    SaisBorMyAfterImport (Sender,FileGuid,Cancel)
  else  if TForm(Sender) is TFSaisie then
    SaisieMyAfterImport (Sender,FileGuid,Cancel)
  else  if (Sender is TFVierge) and (TForm(Sender).Name = 'CPSAISIEPIECE') then
    SaisieParamMyAfterImport (Sender,FileGuid,Cancel)
  else if (ctxPCL in V_PGI.PGIContexte) then  // CA - 16/11/2006 - Intégration en GED hors saisie uniquement en mode PCL
    // ---- Insertion classique dans la GED avec boite de dialogue ----
    begin
     {$IFNDEF TT}
      // Propose le rangement dans la GED
      ParGed.SDocGUID := '';
      ParGed.NoDossier := V_PGI.NoDossier;
      ParGed.CodeGed := 'A';
      ParGed.Annee := FormatDateTime('yyyy', VH^.Encours.Fin);
      ParGed.Mois := FormatDateTime('mm', VH^.Encours.Fin);
      // FileId est le n° de fichier obtenu par la GED suite à l'insertion
      ParGed.SFileGUID := FileGuid;
      {$endif}
      // Description par défaut du document à archiver...
      if Sender is TForm then ParGed.DefName := TForm(Sender).Caption;

      Application.BringToFront;

      if ((not (ctxPCL in V_PGI.PGIContexte )) or ((ctxPCL in V_PGI.PGIContexte) and (JaileDroitConceptBureau (187315)))) then
      begin
        if ShowNewDocument(ParGed)='##NONE##' then
          // Fichier refusé, suppression dans la GED
          V_GedFiles.Erase(FileGuid);
      end
      else
      begin
        if PgiAsk ('Vous n''avez pas le droit d''insérer un document dans la GED'#10' Voulez-vous enregistrer ce document sur disque?') = mrYes then
        with TSaveDialog.Create (nil) do
        begin
          Options    := [ofOverwritePrompt];
          FileName   := GetFileNameGed (FileGUID);
          Nomfichier := ExtractFileExt (FileName);
          Filter     := 'Fichiers *' + NomFichier +'|*' + NomFichier + '|Tous les fichiers (*.*)|*.*';
          if Execute = TRUE then
          if V_GedFiles.Extract (FileName, FileGUID) = TRUE then
            PgiInfo ('Le document a été enregistré: ' + FileName)
          else
            PgiInfo ('Impossible d''enregistrer le document sous: ' + FileName);
            Free;
        end;
        V_GedFiles.Erase(FileGUID);
      end;
    end;
end ;
{$ENDIF}

function MyAglMailForm ( var Sujet : HString; var A, CC: String ; Body: HTStringList ; var Files: String ;
ZoneTechDataVisible: boolean=True) : TResultMailForm ;
begin;
    Result := rmfBad;
    if  ShowNewMessage('', '', '', A,V_PGI.NoDossier,Files) then
    Result := rmfOk;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Indique si les modules de révision sont accessibles.
Suite ........ : MODIF PACK AVANCE
Mots clefs ... : PACK AVANCE;REVISION
*****************************************************************}
Function OkRevis : boolean ;
begin
  Result := V_PGI.Controleur ;
  if Not Result then
    PGIBox('Fonction interdite : vous devez être réviseur pour accéder aux fonctions de révision','Comptabilité') ;
end;




{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 04/03/2004
Modifié le ... :   /  /
Description .. : - LG -04/03/2004 - test pour savoir s'il reste un verrou sur le
Suite ........ : lettrage en saisie avant de lancer la requete
Mots clefs ... :
*****************************************************************}
procedure ControleLettrageEnSaisie ;
var
 lQ : TQuery ;
begin
 if not CEstBloqueLettrage then exit ;
 lQ:=nil;
 try
  lQ := OpenSQL('select E_DATECOMPTABLE,E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE from ECRITURE where E_QUALIFPIECE="L" ' , true) ;
  if not lQ.EOF then
   LanceSaisieFolioL(lQ,taModif) ;
 finally
  Ferme(lQ) ;
 end ;
end;


function AffichageEcranAccueil : boolean;
var stAccueil : string;
    bRetour : boolean;
begin
  if (VH^.Revision.Plan <> '') and
     (ctxPCL in V_PGI.PGIContexte) then
  begin // On gère la révision
    if JaileRoleCompta( RcExpert ) then
      stAccueil := GetParamSocSecur('SO_CPECRANACCUEILN4', '', True)
    else
      if JaiLeRoleCompta( RcSuperviseur ) then
        stAccueil := GetParamSocSecur('SO_CPECRANACCUEILN3', '', True)
      else
        if JaiLeRoleCompta( RcReviseur ) then
          stAccueil := GetParamSocSecur('SO_CPECRANACCUEILN2', '', True)
        else
          stAccueil := GetParamSocSecur('SO_CPECRANACCUEIL', '', True);

    bRetour   := LanceEcranAccueil( stAccueil );
  end
  else
  begin
    stAccueil := GetParamSocSecur('SO_CPECRANACCUEIL', '', True);
    bRetour   := LanceEcranAccueil( stAccueil );
  end;

  Result := bRetour;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function LanceEcranAccueil( vstAccueil : string ) : Boolean;
begin
  Result := False;
  if (vstAccueil = '') or (vstAccueil = 'CMO') then
  begin
    Result := True;
    Exit;
  end;

  if VH^.OkModRIC then
  begin
    if ((vstAccueil = 'COC') and (JaiLeDroitTag( 52110))) then
      CPLanceFiche_CPCONSGENE() // Consultation des comptes
      else
        if ((vstAccueil = 'COE') and (JaiLeDroitTag(7602))) then
          OperationsSurComptes('','','') // Consultation des écritures
        else
          if ((vstAccueil = 'ISA') and (JaiLeDroitTag(52010))) then
            CPLanceFiche_JournalCentralisateur
          else
            if (vstAccueil = 'PTV') and (JaiLeDroitTag(52152) or JaiLeDroitTag(52153)) then
              CPLanceFiche_CPREVPROGTRAVAIL( VH^.EnCours.Code ); // Programme de travail
  end;            
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2002
Modifié le ... : 05/12/2006
Description .. : 
Mots clefs ... : 
*****************************************************************}
function OkDispatchPCL10  : Boolean;
Var i : Integer ;
    St,Nom,Value : String ;
BEGIN
  Result := True;

  if ctxPCL in V_PGI.PGIContexte then
  begin
    for i:=1 to ParamCount do
    begin
      St:=ParamStr(i) ; Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
      Value:=UpperCase(Trim(St)) ;

      if Nom='/ETEBAC' then
       begin
        CPLanceFiche_SaisieTresorerie( 'DP');
        Result := False;
        Exit;
       end;

      if Nom='/ETATSCHAINES' then
      begin
        CPEtatsChainesAuto('CP');
        Result := False;
        Exit;
      end;

      // Traitement par lot de l'alignement des standards
      if Nom = '/ALIGNESTD' then
      begin
        CPAlignementMultiDossier( Value );
        Result := False;
        Exit;
      end;
    end;

    // Gestion du standard utilisé dans le dossier
    MiseAJourPlanRefence;

    {$IFDEF NETEXPERT}
    // Import automatique (Récupération de données)
    ImportAutomatique1;
    {$ENDIF}
  end ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe AYEL
Créé le ...... : 02/11/2007
Modifié le ... : 02/11/2007
Description .. : Lancement d'un état libre
Suite ........ : Fusion de LanceEtatLibreS5 avec LanceEtatLibreS7
Mots clefs ... :
*****************************************************************}
Procedure LanceEtatLibre ( NatureEtat : String ) ;
Var
  CodeD : String;
begin
  CodeD:=Choisir('Choix d''un état libre','MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="'+NatureEtat+'" AND MO_MENU="X"','', False, False, 999999112, '');
  if CodeD<>'' then LanceEtat('E',NatureEtat,CodeD,True,False,False,Nil,'','',False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe AYEL
Créé le ...... : 02/11/2007
Modifié le ... : 02/11/2007
Description .. : Lancement d'une liste libre
Suite ........ : 
Mots clefs ... :
*****************************************************************}
Procedure LanceListeLibre ( NatureEtat : String ) ;
Var
  CodeD : String;
begin
  CodeD:=Choisir('Choix d''une liste libre','MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="'+NatureEtat+'" AND MO_MENU="X" AND MO_EXPORT="X"','', False, False, 999999112, '');
  if CodeD<>'' then LanceEtat('E',NatureEtat,CodeD,True,True,False,Nil,'','',False);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Indique si les modules de révision sont accessibles.
Suite ........ : MODIF PACK AVANCE
Mots clefs ... : PACK AVANCE;REVISION
*****************************************************************}
Function GetStRevis : String ;
BEGIN
  Result := '' ;
  if VH^.DateRevision>IDate1900 then
    Result:='Révision au ' +FormatDateTime('c',VH^.DateRevision) ;
END ;

////////////////////////////////////////////////////////////////////////////////

Procedure VerifInstall (BobPrefixe,BobSuffixe,NomChampParamSoc : String) ;
Type terrbob = (bobOK, bobPasla, bobLaMaisPasIntegre, bobPasLaMaisIntegre, bobAInteger) ;
Var Q : tQuery ;
    BobName,St1 : String ;
    FindIt : Boolean ;
    datecreation : tDateTime ;
    errb : terrbob ;
BEGIN
//BobName:='CIS50751M60.BOB' ;
// SO_I_FISACTIVTPTF
{$IFDEF EAGLCLIENT}
AvertirCacheServer('YMYBOBS') ;
AvertirCacheServer('PARAMSOC') ;
{$ENDIF}
If CS3 Then BobName:='CCS5'+BobPrefixe+BobSuffixe
       Else BobName:='CCS5'+BobPrefixe+BobSuffixe ;
ErrB:=BobOK ;
If BobName='' Then Exit ;
{$IFDEF IMMOSIC}
Exit ;
{$ENDIF}
If NomChampParamSoc='' Then Exit ;

Q := OpenSql('SELECT SOC_NOM FROM PARAMSOC WHERE SOC_NOM ="'+NomChampParamSoc+'" ',True, 1);
If Q.Eof Then ErrB:=bobPasLaMaisIntegre ;
Ferme(Q) ;

If ErrB=BobOK Then Exit ;

Q := OpenSql('SELECT * FROM YMYBOBS WHERE YB_BOBNAME="' + BobName + '"',True, 1);
DateCreation:=iDate1900 ;
if not q.eof then datecreation := q.findfield('YB_BOBDATECREAT').asdatetime;
FindIt := not Q.Eof;
Ferme(Q);
If FindIt Then
  BEGIN
  ErrB:=bobLaMaisPasIntegre ;
  END Else
  BEGIN
  ErrB:=bobPasla ;
  END ;

Case ErrB Of
  bobPasla : BEGIN
             St1 := 'UN FICHIER SYSTEME ('+BobName+'.BOB) N''A PAS ETE INTEGRE. ';
             St1 := St1 +#13#10 +#13#10 +'CELA PEUT DECLENCHER DES DYSFONCTIONNEMENTS.';
             St1 := St1 +#13#10 +#13#10 +'NOUS VOUS CONSEILLONS D''APPELER LE SUPPORT TELEPHONIQUE CEGID.';
             PGIInfo(St1,'AVERTISSEMENT');
             END ;
  bobLaMaisPasIntegre : BEGIN
                        St1 := 'UN FICHIER SYSTEME ('+BobName+'.BOB intégré le '+DateToStr(DateCreation)+ ') N''EST PAS CORRECT. ';
                        St1 := St1 +#13#10 +#13#10 +'CELA PEUT DECLENCHER DES DYSFONCTIONNEMENTS.';
                        St1 := St1 +#13#10 +#13#10 +'NOUS VOUS CONSEILLONS D''APPELER LE SUPPORT TELEPHONIQUE CEGID.';
                        PGIInfo(St1,'PROBLEME DETECTE ! ');
                        END ;
  END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 19/10/2006
Modifié le ... :   /  /
Description .. : - LG - 19/10/2006 - FB 18592 - 18998 - branchement de la
Suite ........ : ged en eAgl
Mots clefs ... :
*****************************************************************}
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
var
    stAExclureFavoris,st, lst : string;
    i : integer ;
  {$IFDEF NETEXPERT}
    NoSEqNet            : integer;
    Retour              : string;
    OkAuto              : Boolean;
    Q                   : TQuery;
  {$ENDIF}
BEGIN
  // MODIF PACK AVANCE
  if not V_PGI.OutLook then // Pour empêcher la mise en place inside
    PRien := nil ;          //  en mode non outlook

  case Num of
    10 : begin
      {$IFDEF SCANGED}
      if V_PGI.RunFromLanceur then
        InitializeGedFiles(V_PGI.DefaultSectionDbName, MyAfterImport)
      else
        InitializeGedFiles(V_PGI.DbName, MyAfterImport);
      V_PGI.GedActive := (VH^.OkModGed);
      {$ENDIF}
      ChargeHalleyUser ;

      ControleLettrageEnSaisie ;
      { Affichage de l'écran d'accueil }
      if not (ctxPCL In V_PGI.PGIContexte) then
        if AffichageEcranAccueil or GetParamSocSecur('SO_CPSPECIFIQUES', False) then
          RetourForce := True;
      UpdateComboslookUp ;

      TripoteStatus(GetStRevis);

      if not OkDispatchPCL10 then
      begin
        // Fermeture de l'application
        Application.ProcessMessages;
        FMenuG.ForceClose := True ;
        FMenuG.Close ;
        Exit;
      end;

      // Route sur la messagerie PGI
      ChargeNoDossier; {JP 31/07/06 : Chargement de NoDossier en PGE}
      if VH^.OkModMessagerie then OnAglMailForm := MyAglMailForm;
      if ( ctxPCL in V_PGI.PGIContexte ) then
      begin

      { Lien Publifi }
      V_PGI.EnablePubliFi := TRUE;
      V_PGI.OnPubliFiBeforePrint := lkPubliFi.BeforePubliFiPrint;
      V_PGI.OnPubliFiAfterPrint := lkPubliFi.AfterPubliFiPrint;

      {$IFDEF NETEXPERT} // lien Netexpert
        // Reprise Net expert ajout me 7/09/2004
        if GetFlagAppli ('CCS5.EXE') and (InterrogeCorbeil ('COMPTA', 'TRA', FALSE) <> nil) then
        begin
           V_PGI.ZoomOLE := TRUE;
           if GetParamSocSecur('SO_CPMODESYNCHRO', FALSE, TRUE) = FALSE then
           begin
                SetParamSoc ('SO_CPLIENGAMME','S1');
                SetParamsoc ('SO_CPMODESYNCHRO', TRUE);
                SetParamsoc ('SO_FLAGSYNCHRO', 'SYN');
           end;
           OkAuto := GetParamSocSecur ('SO_CPSYNCHROAUTOMATIQUE', FALSE);
           if not OkAuto then
           begin
             Q := OpenSQL ('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CPSYNCHROAUTOMATIQUE"', TRUE);
             if not Q.EOF then
             begin
                  if Q.FindField ('SOC_DATA').asstring = '1' then
                  begin
                      PGIInfo ('Par défaut la synchronisation des données sera réalisée automatiquement avec '+RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), False) +
                      #10#13 +' Si vous souhaitez une synchronisation manuelle, veuillez décocher l''option synchronisation automatique'+
                      #10#13 + ' dans les paramètres sociétés, échanges ');
                        SetParamSoc ('SO_CPSYNCHROAUTOMATIQUE', TRUE);
                  end;
             end;
             Ferme (Q);
           end;
           Retour := AGlLanceFiche ('CP','NETLISTEFICHIER','','','COMPTA;TRA');
           if Retour = 'FALSE' then
           begin
              Application.ProcessMessages;
              FMenuG.ForceClose := True ;
              FMenuG.Close ;
              Exit;
           end;
           V_PGI.ZoomOLE := FALSE;
        end;
        ControlEnvoiBL;
      {$ENDIF}
      end;

      FMDispS5.InitBStatutLiasse;
      FMDispS5.InitBStatutRevision;

      // Alignement automatique de la révision si modification du plan
      ControlePlanRevision;

      // ajout me fiche 20546
      if (_EstBloque('NRTOUTSEUL', True)) then
      begin
        PGIInfo ('Attention, un traitement exclusif est en cours !');
        Application.ProcessMessages;
        FMenuG.ForceClose := True ;
        FMenuG.Close ;
        Exit;
      end;

    end;

    11 : ;// Après déconnexion

    12 : BEGIN
      MenuPCL := False ;
      //Récupération du paramètre / MENUPCL en PGE
      if not (ctxPCL in V_PGI.PGIContexte ) then
      begin
        for i:=1 to ParamCount do
        begin
          St:=ParamStr(i) ;
          if UpperCase(Trim(ReadTokenPipe(St,'='))) ='/MENUPCL' then
            MenuPCL := True ;
        end ;
      end ;

      // Interdiction de lancer en direct
      if (Not V_PGI.RunFromLanceur) and (V_PGI.ModePCL='1') then
      begin
        FMenuG.FermeSoc;
        FMenuG.Quitter;
        exit;
      end;
      InitSerialisation;
      if V_PGI.ModePCL='1' then
      begin
        V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
        If CS3 Then RenseignelaSerie(ExeCCS3) Else RenseignelaSerie(ExeCCS5) ;
      end;
      if ctxPCL in V_PGI.PGIContexte then
      BEGIN
        {$IFNDEF EAGLCLIENT}
        { CA - 04/06/2007 - Le bouton de changement d'utilisateur n'existe pas en CWAS }
        FMenuG.bUser.Hint := TraduireMemoire('Exercice de référence');
        FMenuG.bUser.Onclick := FMDispS5.ExercClick;
        {$ENDIF}
        V_PGI.CodeProduit:='0011';
        V_PGI.QRPDF := True;
        // Activation du comptage des impressions
        V_PGI.CompterLesimpressions:=true;
        V_PGI.ProcCompteImpression := CompteImpression ;
      END ;
    END ;

    13 :  begin
          { Pour ne pas déclencher les ChargeMenuPop }
          FMenuG.Outlook.OnChangeActiveGroup := nil;
          if ctxPCL in V_PGI.PGIContexte then begin
          InitialiserTableCompta;
          {$IFDEF SCANGED}
          FinalizeGedFiles;
          {$ENDIF}
          EnvoiBL;
        end;
    end;

    15 : ;

    16 : begin
          {JP 06/07/07 : FQ 18524 : Nouvelle gestion du menu specifique : on le met par défaut,
                         et s'il le faut on le cache}
          if not GetParamSocSecur('SO_CPSPECIFIQUES', False, True) then begin
            for i := 0 to 29 do
              if FMenuG.FModuleNum[i] = 324 then FMenuG.FModuleNum[i] := - 1;
          end;
         // SBO 03/07/2007 : ajout des code specif compta en base
         V_PGI.Specif := V_PGI.Specif + ',' + GetParamSocSecur('SO_CPCODESPECIF', '', True) ;
        //if (not (ctxPCL in V_PGI.PGIContexte)) then
          //ExportS7ImportS5Confidentialite;
(*        // PGI_IMPORT_BOB('CCS5');
        { Importation des bobs CCS5 }
        BOB_IMPORT_PCL('CCS5');
        { Importation des BOBS communes en Entreprise }
        if not (ctxPCL in V_PGI.PGIContexte) then BOB_IMPORT_PCL('COMM',True);
*)
        St := 'CCS5';
        lSt := 'Paramétrage Comptabilité';
        if not (ctxPCL in V_PGI.PGIContexte) then begin
          St := St+';COMM';
          lSt := lSt + ';Paramétrage Commun';
        end;
// Modif GP : plus d'import auto en mode CWAS.
//     BOB_IMPORT_PCL_STD (St,lSt);
{$IFDEF EAGLCLIENT}
{$IFDEF ENTREPRISE}
        VerifInstall ('0851M','902','SO_CP_NEWRAPPRO') ;
{$ELSE}
        BOB_IMPORT_PCL_STD (St,lSt);
{$ENDIF}
{$ELSE}
        BOB_IMPORT_PCL_STD (St,lSt);
{$ENDIF}

    end;

    100 : BEGIN
            if (ctxPCL in V_PGI.PGIContexte) then
            begin
              { vérif du mot de passe lors d'une connexion automatique sur un dossier PCL protégé }
              If Not CheckPwdDossier( V_PGI.NoDossier, RecupPwdDossierFromLgcde) then
              begin
                FMenuG.FermeSoc;
                FMenuG.Close;
                exit;
              end;
              { On force le mode CascadeForms }
              ForceCascadeForms;
              { Lancement de l'assistant si le dossier n'a pas été initialisé }
              if not LanceAssistantComptaPCL(True) then
              begin
                FMenuG.ForceClose := True;
                FMenuG.Close;
                Exit;
              end;
              { Affichage de l'écran d'accueil }
              if AffichageEcranAccueil then RetourForce := True;
            end
            else if ((ctxCompta in V_PGI.PGIContexte) and (V_PGI.RunFromLanceur)) then SetFlagAppli('CCS5.EXE',True);
    end;


    4006 : AGLLanceFiche ('YY','YYDEFPRT','','','TYPE=E');

//======================
//==== SPECIFIQUES  ====
//======================
{$IFNDEF EAGLCLIENT}

{$IFDEF SECART}
    7010 : AGLLanceFiche('CP', 'ECSAISIE', '', '', '') ;
    7011 : AGLLanceFiche('CP', 'ECCOMPTES', '', '', '') ;
    7012 : AGLLanceFiche('CP', 'ECCALCUL', '', '', '') ;
{$ENDIF}
{$ENDIF}

//========================
//==== Ecritures (9)  ====
//========================
    // Courantes
    {$IFDEF TT}
    7244 : CPLanceFiche_JournalCentralisateur;
    {$ELSE}
    7244 : MultiCritereMvt(taCreat,'N',False) ;       // Création
    {$ENDIF}
    7241 : CPLanceFiche_MulGuide('','','SELECT;') ;// LG - 12/09/2007 - suppression de la fct SelectGuide
    7256 : MultiCritereMvt(taConsult,'N',False) ;     // Consultation

    7259 : if ctxPcl in V_Pgi.PgiContexte then
             CPLanceFiche_CPRechercheEcr(False)
           else
             MultiCritereMvt(taModif, 'N', False);    // Modification

    7262 : DetruitEcritures('N',False) ;              // Suppression Ecritures Pieces
    7263 : DetruitEcritures('N',True) ;               // Suppression Ecritures Bordereaux
    7005 : SaisieFolio(taModif) ;
    52116 : CPLanceFiche_CPMULCUTOFF('') ;
    7260 : ModifSerieTableLibreEcr('E',[tpReel]) ; // MODIF PACK AVANCE {JP 10/06/05 : migration en eAGL}
    7247 : PrepareSaisTres ;
    7245 : PrepareSaisTresEffet ;
    7264 :  if ctxPCL in V_PGI.PGIContexte Then CPLanceFiche_VisuTresorerie  //LG* 06/05/2002
                                           Else CCLanceFiche_ModifEntPie;

    7268 : CPLanceFiche_CPBROUILLARDECR('N'); {FP 28/12/2005} //Brouillard('N') ;

    // Modules SAISIE paramètrables
    9105 : SaisiePieceCompta ;                        // Création Normale
    9106 : SaisiePieceCompta('I') ;                   // Création IFRS
    9107 : SaisiePieceCompta('S') ;                   // Création Simulation
    9108 : SaisiePieceCompta('U') ;                   // Création Situation
    9109 : SaisiePieceCompta('P') ;                   // Création Prévision
    9110 : SaisiePieceCompta('R') ;                   // Création Révision

    9115 : MultiCritereMvt(taConsult,'SPN',False) ;   // Consultation Normale
    9116 : MultiCritereMvt(taConsult,'SPI',False) ;   // Consultation IFRS
    9117 : MultiCritereMvt(taConsult,'SPS',False) ;   // Consultation Simulation
    9118 : MultiCritereMvt(taConsult,'SPU',False) ;   // Consultation Situation
    9119 : MultiCritereMvt(taConsult,'SPP',False) ;   // Consultation Révision
    9120 : MultiCritereMvt(taConsult,'SPR',False) ;   // Consultation Révision

    9125 : MultiCritereMvt(taModif,'SPN',False) ;     // Modification Normale
    9126 : MultiCritereMvt(taModif,'SPI',False) ;     // Modification IFRS
    9127 : MultiCritereMvt(taModif,'SPS',False) ;     // Modification Simulation
    9128 : MultiCritereMvt(taModif,'SPU',False) ;     // Modification Situation
    9129 : MultiCritereMvt(taModif,'SPP',False) ;     // Modification Prévision
    9130 : MultiCritereMvt(taModif,'SPR',False) ;     // Modification Révision

    // Modules SAISIE IFRS
    9200 : MultiCritereMvt(taCreat,'I',False) ;       // Création
    9210 : MultiCritereMvt(taConsult,'I',False) ;     // Consultation
    9220 : MultiCritereMvt(taModif,'I',False) ;	      // Modification
    9230 : DetruitEcritures('I',False) ;              // Suppression Ecritures Pieces IFRS
    9240 : CPLanceFiche_TraitementIAS14('', '', '') ; // Reventilation IAS14

    // Simulation
    7277 : MultiCritereMvt(taCreat,'S',False) ;
    7274 : MultiCritereMvt(taConsult,'S',False) ;
    7280 : MultiCritereMvt(taModif,'S',False) ;
    7283 : DetruitEcritures('S',False) ;
    7282 : ModifSerieTableLibreEcr('E',[tpSim]) ; // MODIF PACK AVANCE {JP 14/06/05 : Migration eAgl}
    7281 : CPLanceFiche_ValidePiece('', '', 'S');
    7289 : CPLanceFiche_CPBROUILLARDECR('S'); {FP 28/12/2005} //Brouillard('S') ;

    // Situation
    7298 : MultiCritereMvt(taCreat,'U',False) ;
    7295 : MultiCritereMvt(taConsult,'U',False) ;
    7301 : MultiCritereMvt(taModif,'U',False) ;
    7302 : CPLanceFiche_ValidePiece('', '', 'U');
    7304 : DetruitEcritures('U',False) ;
    7305 : ModifSerieTableLibreEcr('E',[tpSitu]) ; // MODIF PACK AVANCE {JP 14/06/05 : Migration eAgl}

    // Prévisions
    7310 : MultiCritereMvt(taConsult,'P',False) ;   // MODIF PACK AVANCE
    7313 : MultiCritereMvt(taCreat,'P',False) ;     // MODIF PACK AVANCE
    7316 : MultiCritereMvt(taModif,'P',False) ;     // MODIF PACK AVANCE
    7317 : ModifSerieTableLibreEcr('E',[tpPrev]) ;  // MODIF PACK AVANCE {JP 14/06/05 : Migration eAgl}
    7319 : DetruitEcritures('P',False) ;            // MODIF PACK AVANCE

    // Révision
    7662 : if OkRevis then CPLanceFiche_ValidePiece('', '', 'R');           // MODIF PACK AVANCE
    7649 : if OkRevis then MultiCritereMvt(taCreat,'R',False) ;     // MODIF PACK AVANCE
    7658 : if OkRevis then MultiCritereMvt(taConsult,'R',False) ;   // MODIF PACK AVANCE
    7661 : if OkRevis then MultiCritereMvt(taModif,'R',False) ;     // MODIF PACK AVANCE
    7664 : if OkRevis then DetruitEcritures('R',False) ;            // MODIF PACK AVANCE
    7682 : if OkRevis then LanceRevision ;      // MODIF PACK AVANCE
    7670 : if OkRevis then CPLanceFiche_CPBROUILLARDECR('R'); {FP 28/12/2005} //Brouillard('R') ;  // MODIF PACK AVANCE

    // Abonnements
    7328 : GenereAbonnements(True) ;
    7331 : GenereAbonnements(False) ;
    7337 : ParamAbonnement(TRUE,'',taModif) ;
    7340 : ReconductionAbonnements ;
    7346 : ListeAbonnements ;

{b Thl 02/05/2006}
    // Engagements
    7348 : MultiCritereMvt(taCreat,'SPp',False) ;
    7349 : MultiCritereMvt(taConsult,'SPp',False) ;
    7350 : MultiCritereMvt(taModif,'SPp',False) ;
    7351 : DetruitEcritures('p',False) ;
    7356 : CPLanceFiche_CPGESTENG;  // thl : en commentaire si Cwas
    7357 : CPLanceFiche_CPRAPAUTOENG;
    7358,  {Lek pourquoi eCCS5 n'est pas sur 7359???}
    7359 : CPLanceFiche_CPSuiReaEng;
{e Thl 02/05/2006}
    // Analytiques
    7113 : CCLanceFiche_MULGeneraux('SERIEANA;7115000;1');
    7367 : MultiCritereAna(taCreat) ;
    7364 : MultiCritereAna(taConsult) ;
    7370 : MultiCritereAna(taModif) ;
    7382 : CPLanceFiche_ReventilAna;
    7379 : VidageInterSections ;             // MODIF PACK AVANCE
    7383 : CPLanceFiche_LoiVentilMul('', '', 'CLV_LOITYPE=LOI');
    7384 : CPLanceFiche_LoiVentilMul('', '', 'CLV_LOITYPE=SEC');
    7371 : ModifSerieTableLibreEcr('A',[]) ; // MODIF PACK AVANCE {JP 14/06/05 : Migration eAgl}
    7373 : CPLanceFicheMulSupprAna;  //    7373 : DetruitAnalytiques ;
    7381 : CC_LanceFicheReimpAna; // CA- 26/04/2007 - Fusion des réimputations analytiques Web Access et 2 Tiers
    7385 : CPLanceFiche_CPBROUILLARDANA;  {FP 28/12/2005}

//====================
//==== Tiers (11) ====
//====================
    // Encaissements
{$IFDEF EAGLCLIENT}
{$ELSE}
    7496 : EncaisseDecaisse(True,'','',True,False,tslAucun) ;
    7497 : EncDecChequeTraiteBOR(tslTraite) ;
{$ENDIF}

    // Décaissements
{$IFDEF EAGLCLIENT}

{$ELSE}
    74999: EncaisseDecaisse(False,'','',True,True,tslAucun) ;
    74991: DecaisseCircuit(1) ;
    74992: DecaisseCircuit(2) ;
    74993: DecaisseCircuit(3) ;
    74994: DecaisseCircuit(4) ;
    74995: DecaisseCircuit(5) ;
    7501 : EncDecChequeTraiteBOR(tslCheque) ;
    7502 : EncDecChequeTraiteBOR(tslBOR) ;
{$ENDIF}
    74997: LanceBonAPayer ;

    // Lettrage
    7508 : LanceLettrage;
    7511 : LanceDeLettrage;
    11440: LettrageParCode;
    7514 : RapprochementAuto('','');
    7520 : RegulLettrage(True,False) ;
    7523 : RegulLettrage(False,False) ;
    7524 : RegulLettrage(False,True) ;
{$IFDEF EAGLCLIENT}
{$ELSE}
    7529 : ModifEcheances ;
{$ENDIF}

    // Relances clients
    7565 : RelanceClient(True,True) ;    // relance manuelle sur traites
    7566 : RelanceClient(True,False) ;   // relance manuelle sur autres modes paiement
    7567 : RelanceClient(False,True) ;   // relance auto sur traites
    7568 : RelanceClient(False,False) ;  // relance auto sur autres modes paiement
    7583 : CalculScoring ;

    //YMO 03/07/2006 FQ 17850 Menu Estimation des coûts financiers dispo en CWAS
    7598 : LanceCoutTiers('',0) ;

    // Editions de suivi
    7547 : CPLanceFiche_BalanceAgee;
    7556 : CPLanceFiche_BalanceVentilee;  // FQ 14149 - 21/01/2005
    7560 : CPLanceFiche_CPGLAUXI(''); // GLAuxiliaireL ; // FQ 13981 / SBO 29/09/2004
    7541 : AGLLanceFiche('CP', 'EPECHEANCIER','','ECH','ECH');
    7535 : CPLanceFiche_JustSold;//JustSolde ; Lek remplace 190106
    7550 : CPLanceFiche_CPTGLAge;// GLivreAge ; Lek remplace 190106
    7559 : CPLanceFiche_CPTGLVentil;// GLVentile ; Lek remplace 190106

    // Tiers payeurs
    7231 : AGLLanceFiche('CP', 'EPTIERSP', '', 'TFP', 'TPF') ; // Tiers facturé par payeur
    7232 : AGLLanceFiche('CP', 'EPTIERSP', '', 'TFA', 'TPF') ; // Tiers facturé avec payeur
{.  $IFDEF EAGLCLIENT Lek 190106}
    7233 : CPLanceFiche_CPTPAYEURFACTURE;{Lek&FP 071005 BrouillardTP ;}
    7234 : CPLanceFiche_BalAgeeTPayeur;  {Lek&Fréd 111005}
    7235 : CPLanceFiche_BalVentilTPayeur; {Lek&Fréd 241005 BalVentil}
    7236 : CPLanceFiche_CPTPayeurGLAge;   {Lek&FP 211005 GLivreAgeTP ;}
    7237 : CPLanceFiche_CPTPAYEURGLVENTIL; {Lek&FP 241005 GLivreVenil}

    7239 : CPLanceFiche_CPGeneEcrTP('', '', 'S'); //YMO 12/06 Génération manuelle sur TP

{  $ELSE
    7233 : BrouillardTP ;
    7234 : BalanceAge2TP ;
    7236 : GLivreAgeTP ;
    7235 : BalVentileTP ;
    7237 : GLVentileTP ;
 $ENDIF}

    //Emissions
    7589 : ReleveCompte ;   // Relevés de comptes
    7592 : ReleveFacture ;  // Relevés de factures
    7585 : ModifEcheMP;

    7599 : ExportCFONBBatch(False,True,tslAucun) ;
    7586 : ExportCFONBBatch(True,True,tslAucun) ;
    7595 : ExportCFONBBatch(True,False,tslAucun) ;
    7596 : ExportCFONBBatch(False,False,tslTraite) ;
    7597 : ExportCFONBBatch(False,False,tslBOR) ;
{$IFDEF EAGLCLIENT}
{$ELSE}
    7594 : AGLLanceFiche('CP', 'RLVEMISSION', '', '', '') ;
{$ENDIF}

//==== Editions (13) ====
    // Journaux
    {JP 23/06/05 : Uniformisation des journaux PGE - PCL
                   Le menu est renommé en journal des écritures et non plus journal divisionnaire}
    7394 : CPLanceFiche_CPJALECR ; // 26/11/2002 nv journal des ecritures
    7409 : CPLanceFiche_CPJALCPTGENE;  {FP 28/12/2005} //JalCpteGe  ;

    // Grand-livres
    7415 : CPLanceFiche_CPGLGENE('');
    7418 : CPLanceFiche_CPGLAUXI('');
    7419 : CPLanceFiche_EtatGrandLivreSurTables; 	// Grand Livre sur table libre
    7427 : CPLanceFiche_CPGLGENEPARAUXI;
    7430 : CPLanceFiche_CPGLAUXIPARGENE;
    7421 : CPLanceFiche_CPGLANA ;
    7436: CPLanceFiche_CPGLGENEPARANA;
    7439: CPLanceFiche_CPGLANAPARGENE;
    7417 : CPLanceFiche_CPGLGENEPARQTE;
    7420 : CPLanceFiche_CPGLGESTION;

    // Balances
    7445 : CPLanceFiche_BalanceGeneral; 												// Balance générale eAGL
    7448 : CPLanceFiche_BalanceAuxiliaire; 			 								// Balance auxiliaire eAGL
    7429 : CPLanceFiche_EtatBalanceSurTables; 									// Balance sur table libre
    7451 : CPLanceFiche_BalanceAnalytique ;                     // Balance analytique
    7457 : CPLanceFiche_BalanceGenAuxi ;
    7460 : CPLanceFiche_BalanceAuxiGen ;
    7466 : CPLanceFiche_BalanceGenAnal ;
    7469 : CPLanceFiche_BalanceAnalGen ;

    // SBO 16/11/2005 : remplacement ancien cumul pour nouveau
    7478 : CPLanceFiche_CumulPeriodique( 'QRS1CUMGEN' );
    7481 : CPLanceFiche_CumulPeriodique( 'QRS1CUMAUX' );
    7484 : CPLanceFiche_CumulPeriodique( 'QRS1CUMSEC' );
    // BVE 29.05.07
    52603 : CPLanceFiche_CumulPeriodique( 'QRS1CUMGEN' );
    52604 : CPLanceFiche_CumulPeriodique( 'QRS1CUMAUX' );
    52605 : CPLanceFiche_CumulPeriodique( 'QRS1CUMSEC' );
    // Etats de synthèses
    7440 : EtatPCL (esCR)  ; // Document de synthèse - Résultat
    7441 : EtatPCL (esBIL) ; // Document de synthèse - Actif
    7442 : EtatPCL (esSIG) ; // Document de synthèse - SIG
    7446 : EtatPCL (esCRA) ; // Document de synthèse analytique - Résultat
    7449 : EtatPCL (esBILA) ; // Document de synthèse analytique - Bilan
    7447 : EtatPCL (esSIGA) ; // Document de synthèse analytique - SIG
    7444 : CPLanceFiche_EtatsChaines('CP');
{$IFDEF EAGLCLIENT}
{$ELSE}
    7490 : VisuEditionLegale(FALSE) ; // MODIF PACK AVANCE
{$ENDIF}

    // Editions paramétrables
    7410 : AGLLanceFiche('CP', 'EPJALDIV', '', 'JDP', 'JDP') ; // Journal divisionnaire
    7411 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JCJ', 'JAC') ; // Journal centralisateur par journal
    7412 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JCD', 'JAC') ; // Journal centralisateur par période-journal
    7413 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JPP', 'JPE') ; // Journal périodique par Journal-Période
    7414 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JPJ', 'JPE') ; // Journal périodique par Période-Journal
    7422 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JPN', 'JPE') ; // Journal périodique par Période-Journal-Nature
    7423 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JGE', 'JAL') ; // Journal général
    7424 : AGLLanceFiche('CP', 'EPGLIVRE', '', 'GLP', 'GLP') ; // Grand livre général
    7425 : AGLLanceFiche('CP', 'EPGLIVRE', '', 'GLB', 'GLB') ; // Grand livre auxiliaire
    7416 : AGLLanceFiche('CP', 'EPECHEANCIER','','ECH', 'ECH') ; // Echéancier  ATTENTE -> 'ECP','ECP');
    7426 : AGLLanceFiche('CP', 'EPBGLGEN', '', 'BGE', 'BGE') ; // Balance générale
    7450 : AGLLanceFiche('CP', 'EPBGLGEN', '', 'BAU', 'BAU') ; // Balance auxiliaire
    7428 : AGLLanceFiche('CP', 'EPBGLGEN', '', 'BVE', 'BVE') ; // Balance ventilée
    7454 : AGLLanceFiche('CP', 'EPTABMVT', '', 'JTL', 'JTA') ; // Tableau d'avancement - Mouvements
    7455 : AGLLanceFiche('CP', 'EPTABMVT', '', 'JTM', 'JTA') ; // Tableau d'avancement - Montants
    7434 : if ctxPCL in V_PGI.PGIContexte then EtatPCL (esCR)  // Document de synthèse - Résultat
                                          else AGLLanceFiche('CP', 'EPETASYN', '', 'RCP', 'RES') ; // Document de synthèse - Résultat

    7435 : if ctxPCL in V_PGI.PGIContexte then EtatPCL (esBIL)  // Document de synthèse - Actif
                                          else AGLLanceFiche('CP', 'EPETASYN', '', 'BAF', 'BIL') ; // Document de synthèse - Actif
    7453 : if ctxPCL in V_PGI.PGIContexte then EtatPCL (esSIG)  // Document de synthèse - SIG
                                          else AGLLanceFiche('CP', 'EPETASYN', '', 'SIG', 'SIG') ; // Document de synthèse - SIG

    // Etats libres
    7456 : LanceEtatLibre('UCO') ;
    13720 : LanceListeLibre('UCO');

    // Autres éditions
   52270 : CPLanceFiche_CPCONTROLTVA() ;

// BPY le 21/06/2004 => Restitutions multi sociétés
    // edition multi-dossier
    // BALANCES
    8611 : CPLanceFiche_BalanceGeneral('MULTIDOSSIER');    // Balance générale eAGL
    8612 : CPLanceFiche_BalanceAuxiliaire('MULTIDOSSIER'); // Balance auxiliaire eAGL
    8613 : CPLanceFiche_BalanceAnalytique('MULTIDOSSIER'); // Balance analytique
    8614 : CPLanceFiche_EtatBalanceSurTablesM; 		   // Balance sur table libre
    8615 : CPLanceFiche_BalanceGenAuxi('MULTIDOSSIER');
    8616 : CPLanceFiche_BalanceAuxiGen('MULTIDOSSIER');
    8617 : CPLanceFiche_BalanceGenAnal('MULTIDOSSIER');
    8618 : CPLanceFiche_BalanceAnalGen('MULTIDOSSIER');
    // GRAND LIVRES
    8621 : CPLanceFiche_CPGLGENE('MULTIDOSSIER');
    8622 : CPLanceFiche_CPGLAUXI('MULTIDOSSIER');
    8623 : ;// GLAnal ; => Attente du passage en etat AGL
    8624 : CPLanceFiche_EtatGrandLivreSurTablesM;          // Grand Livre sur table libre
    8625 : CPLanceFiche_CPGLGENEPARAUXI('MULTIDOSSIER');
    8626 : CPLanceFiche_CPGLAUXIPARGENE('MULTIDOSSIER');
    8627 : ;//GLGESE ; //=> Attente du passage en etat AGL
    8628 : ;//GLSEGE ; //=> Attente du passage en etat AGL
// Fin BPY

//==========================
//==== Traitements (16) ====
//==========================
    // Sur comptes
    7602 : OperationsSurComptes('','','');
    52003 : CPLanceFiche_SaisieTresorerie ; // Saisie de trésorerie (PCL)
    52004 : CPLanceFiche_SaisieVrac;        // Saisie en vrac (PCL)
    52010 : CPLanceFiche_JournalCentralisateur;
    {$IFDEF JOHN}
    52103 : CPLanceFiche_CPINTEGRECFONB('');
    {$ENDIF JOHN}
    52110 : CPLanceFiche_CPCONSGENE();      // Consultation des comptes généraux
    52120 : CPLanceFiche_CPCONSAUXI();      // Consultation des comptes auxiliaires
    //Lek 150206
    52130 : CPLanceFiche_ConsSectAna;   //Consultation section analytique
    52140 : CPLanceFiche_ConsEcrAna;    //Consultation écriture avec section analytique

    7261 : Centralisateur(taModif, 'N', FALSE) ;

    52151 : CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.Encours.Code, 2, TaModif, 0);
    52152 : CPLanceFiche_CPRevProgTravail( VH^.EnCours.Code ); // Programme de travail
    52153 : CPLanceFiche_CPRevProgTravail( VH^.EnCours.Code ); // Programme de travail
    52154 : TRic.ImpressionDossierTravail;

    7676 : CC_LanceFicheReimp;
    // Ré-imputations
{$IFDEF EAGLCLIENT}
    7677 : CPLanceFiche_Extourne;
{$ELSE}
    7677 : If Not EstSpecif('51189') Then CPLanceFiche_Extourne
                                     else LanceExtourne ;
    7719 : TLAuxversTLEcr ;
{$ENDIF}

    // anciens specif CERIC 24/03/2006 (dispo en eAGL)
    7725 : GeneMvt ;                     // Ecritures sur opérations à l'achèvement
    7735 : CPLanceFiche_EtatDesMarges ;  // Impression des affaires en cours

    // Pointage
    {JP 05/03/07 : Gestion du nouveau pointage}

    16110 : BEGIN
             If INITNEWPOINTAGE Then
               BEGIN
               AvertirCacheServer('PARAMSOC') ;
               FermeHalley:=TRUE ; 
               END ;
             END ;

    7601 : CPLanceFiche_ImportCEtebac('');
    7603 : CPLanceFiche_MulCEtebac('');
    7605 : CPLanceFiche_MulEexBqLig('', '', '');
    7609 : If OkNewPointage Then
             BEGIN
             if MsgPointageSurTreso then CPLanceFiche_PointageMul(CODENEWPOINTAGE);
             END Else
             BEGIN
             CPLanceFiche_PointageMulV7 ( );
             END ;

    7616 : if MsgPointageSurTreso then CC_LanceFicheEtatPointage;        // Etat de pointage
    {JP 02/10/07 : Ajout du message d'avertissement si pointage sur TRECRITURE}
    7619 : If OkNewPointage Then
             BEGIN
             if MsgPointageSurTreso then CC_LanceFicheEtatRapproDet;          // Etat de rapprochement
             END Else
             BEGIN
             CC_LanceFicheEtatRapproDetV7;          // Etat de rapprochement v2207
             END ;
    7622 : CC_LanceFicheJustifPointage ;      // Justificatif des soldes bancaires
    7631 : CPLanceFiche_ControlCaisse('');

    // Validations
    7691 : ValidationPeriode(TRUE,FALSE);     // Validation par période
    7694 : ValidationPeriode(FALSE,FALSE);    // Dévalidation par période
    7700 : ValidationJournal(TRUE) ;          // Validation par journal
    7703 : ValidationJournal(FALSE) ;         // Dévalidation par journal
    7705 : CPLanceFiche_ValideEcriture('', '', 'V'); // Validation de pièces 10/06
    7707 : CPLanceFiche_ValideEcriture('', '', 'D'); // Dévalidation de pièces 10/06

    7709 : CPLanceFiche_EDTLegal;// LanceEdtLegal ; Lek remplace 190106

    // A-nouveaux
    7718 : OuvertureExo ; // Ouverture d'exercice
    7724 : MultiCritereMvt(taCreat,'N',True) ;
    7727 : MultiCritereMvt(taConsult,'N',True) ;
    // Clôtures
    7736 : CloPer;                          // Clôture périodique
    7737 : AnnuleCloPer;                    // Annulation de clôture périodique
    7742 : begin ClotureComptable(FALSE) ; {RetourForce := True;} end;      // Clôture provisoire //SG6 06/12/2004 FQ 14987
    7745 : begin AnnuleclotureComptable(FALSE); {RetourForce := True;} end; // Annulation de clôture provisoire //SG6 06/12/2004 FQ 14987
    7751 : CPLanceClotureDefinitive;


{$IFDEF EAGLCLIENT}
{$ELSE}   {Autre édition sur traitement Lek 190106}
   77120 : VisuEditionLegale(TRUE) ; // MODIF PACK AVANCE
{$ENDIF}
    7683 : ResultatDeLexo ;
    7757 : CPLanceFiche_CPJALOUVCLO('O');  {FP 28/12/2005} //JalDivisioAno ;
    7760 : CPLanceFiche_CPBALOUVCLO('O');  {FP 28/12/2005} //BalanceAno ;
    7766 : CPLanceFiche_CPJALOUVCLO('F');  {FP 28/12/2005} //JalDivisioClo ;
    7769 : CPLanceFiche_CPBALOUVCLO('F');  {FP 28/12/2005} //BalanceClo ;

    // A-nouveaux analytiques
    7752 : ClotureComptableAna ;      // Clôture analytique
    7753 : AnnuleClotureAnalytique ;  // Annulation clôture analytique
    7728 : MultiCritereAnaANouveau(taCreat);
    7729 : MultiCritereAnaANouveau(taConsult);
    7730 : MultiCritereAnaANouveau(taModif);

    // Balances
    7710 : SaisieBalance('N1A', taModif) ;  // INUTILISE
    7711 : SaisieBalance('N1C', taModif) ; // INUTILISE
    7712 : SaisieBalance('N0A', taModif) ;
    7713 : SaisieBalance('BDS', taCreat) ; // INUTILISE
    7715 : SaisieBalance('BSA', taCreat) ; // INUTILISE

{$IFDEF EAGLCLIENT}
{$ELSE}
    7714 : MultiCritereBds(taModif) ;      // INUTILISE
{$ENDIF}
    6369 : CP_LanceFicheExpEtafi;     // ETAFI Servant Soft

    // Relevés
    {$IFDEF RR}
    181215 : ParamSociete(False,'','SCO_MDETEBAC','',Nil,Nil,Nil,Nil,0);
    181216 : AGLLanceFiche('CP','RLVETEBAC','','',';;;I') ; // intégration des relevés
   {$ELSE}
    7772 : CPLanceFiche_RecepETEBAC3;
   {$ENDIF}

    //7774 : CP_LancePointageAuto;
    //7776 : CPLanceFiche_DPointageMul;
// Ancien rappro :
    7773 : CPLanceFiche_RecupReleve;  // Intégration des relevés
    7774 : CP_LancePointageAuto;
    7776 : CPLanceFiche_DPointageMul;
    7777 : CC_LanceFicheEtatRapproDet;
// Fin Ancien rappro

    // Révision (54710, 54720  uniquement en PGE)
    54710 : AGLLanceFiche('CP', 'CPREVMULPLAN', '', '', '') ;
    54720 : CPLanceFiche_CPREVMULPARAMCYCL('');

    54725 : CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.Encours.Code, 0, TaModif);
    54730 : CPLanceFiche_CPCONTROLECYCLE('');
    54740 : CPLanceFiche_CPCONTROLELIASSE('');
    54750 : CPLanceFiche_CPTOTLIASSE('');

    // FTS
    66141 : CPLanceFiche_CPIMPFICHEXCEL ('FTS');
    66142 : CPLanceFiche_CPIMPFICHEXCEL ('EXL');

    // ICC
    96110 : CPLanceFiche_ICCGENERAUX;      // Comptes courants
    96130 : CPLanceFiche_FicheICCGENERAUX( GetParamSoc('SO_ICCCOMPTECAPITAL') ); // Compte de capital
    96140 : CPLanceFiche_ICCELEMNATIONAUX; // Eléments nationaux

    // CRE
    96210 : AGLLanceFiche('FP', 'FMULEMPRUNT', '', '','');
    96220 : LancerEditions;
    // GCD
    96410 : CPLanceFiche_CPMULGCD ;
    96420 : CPLanceEtatGCD ;
    96240 : EtatRapproCRECompta;

    // DADS2
    96510 : LanceMul_DADS2Tiers;                                //Import des données
    96520 : LanceMul_DADS2Honoraires;                           //Saisie des informations
    96530 : LanceFiche_DADS2Edition;                            //Edition
    96540 : LanceFiche_DADS2Prep;                               //Préparation pour envoi
    96560 : LanceMul_DADS2Envoi('MUL_ENVOISOCIAL', 'DADSB');    //Envoi
    96570 : LanceMul_DADS2Envoi('MUL_CONSULTENVSOC', 'DADSB');  //Consultation
    96580 : LanceMul_DADS2Envoi('MUL_ENVOISOCIAL', 'DADSBP');   //Purge
    96590 : AglLanceFiche('PAY', 'ETABLISSEMENT', '', '', '');

    // EDI
    {$IFDEF NETEXPERT} // lien Netexpert
    96310 :  EchgImp_LanceAssistant ('BAL');
    96320 :  EchgImp_LanceAssistant ('GLJ');
    96315 :  EchgImp_LanceAssistant ('AUT');
       // import
      96321 :  ImportDonnees ('');
       // export
      96322 :  ExportDonnees('', FALSE, 'S5');
      96323 :
      begin
        if (IsDossierNetExpert(V_PGI.NoDossier, NoSEqNet)) then EnvoiBL
        else ExportDonnees('', FALSE, 'S1');
      end;
      96324 :
      begin
        if (IsDossierNetExpert(V_PGI.NoDossier, NoSEqNet)) then
        begin
          // Reprise Net expert ajout me 7/09/2004
          if GetFlagAppli ('CCS5.EXE') and (InterrogeCorbeil ('COMPTA', 'TRA', FALSE) <> nil) then
          begin
             V_PGI.ZoomOLE := TRUE;
             AGlLanceFiche ('CP','NETLISTEFICHIER','','','COMPTA;TRA');
             V_PGI.ZoomOLE := FALSE;
          end;
        end
        else
           ImportDonnees ('');
      end;
      96327 :  CPLanceFiche_JournalDesimports('');

      {$ENDIF NETEXPERT}
{$IFDEF CISXPGI}
    96325 : // Relevés
    begin
              CPLanceFiche_RecupReleveInter ('IMPORT');
    end;
    96326 : // export Relevés
              CPLanceFiche_RecupReleveInter ('EXPORT');
{$ENDIF}

    // Gestion des situations
    7716 : CCLanceFiche_BalanceSituation;
    7717 : CCLanceFiche_BalanceSituation;

    // Analyse de trésorerie
{$IFDEF EAGLCLIENT}
{$ELSE}
    7625 : SuiviTresoParPeriode ; // MODIF PACK AVANCE
{$ENDIF}

//===============================
//==== Reporting et TVA (14) ====
//===============================
    // Reporting
    7787 : MultiCritereRubriqueV2(taModif,False); {Multi-critères des rubriques}
    7790 : SuppressionRubriqueV2(False) ; {Suppression des rubriques}
    7796 : ImpRubrique('', False, False); {Impression des rubriques}
    7775 : ParamTable('CPRUBFAMILLE', taCreat,7775000,PRien, 3, 'Familles de rubriques') ;
    //Sg6 21/01/05 FQ 15288 Gestion Creatio rubrique à partir table libre 
    7815 : CPLanceFiche_CpRubTLib('', '', '') ; //Création des rub à partir des tables libres
{$IFDEF EAGLCLIENT}
    7811 : ParamTable('TTCONSTANTE', taCreat,0, PRien, 3, 'Constantes'); // Constantes

    //SG6 FQ 15288 21/01/05 7815 : AGLLanceFiche('CP', 'CPRUBTLIB', '', '', '') ;{... des tables libres}
{$ELSE}
    // Maquettes d'édition / Etats
    69110 : AGLLanceFiche('CP','STDMAQ','CR','','ACTION=MODIFICATION;CR');   // Compte de résultat (PCL)
    69310 : AGLLanceFiche('CP','STDMAQ','BIL','','ACTION=MODIFICATION;BIL'); // Bilan (PCL)
    69210 : AGLLanceFiche('CP','STDMAQ','SIG','','ACTION=MODIFICATION;SIG'); // SIG (PCL)

    7811 : Constant.Constantes ;
    //SG6 FQ 15288 21/01/05 7815 : AGLLanceFiche('CP', 'MULTABRUB', '', '', '') ;
{$ENDIF}
    {Création de rubriques à partir ...}
    7812 : AGLLanceFiche('CP', 'MULBUDRUB', '', '', '') ;{... des rubriques budgétaires}
    7813 : AGLLanceFiche('CP', 'MULRUPRUB', '', '', '') ;{... des ruptures}
    7814 : AGLLanceFiche('CP', 'MULCORRUB', '', '', '') ;{... des plans de correspondance}
    7802 : ControleRubrique(''); {Contrôle du paramétrage}
    7805 : TotalRubriques ; {Contrôle des montants}

    // Gestion TVA
    7640 : ControleTVAFactures('',0,0,'');  // Contrôle des factures
    7646 : begin   // ajout me CA3
        if ctxPCL in V_PGI.PGIContexte then
//          AGLLanceFiche('CP','MULINFOCA3','','','')
          PGIInfo ('La déclaration de TVA est gérée de manière autonome par rapport à CEGID Comptabilité PGI.#10#13Il convient d''utiliser directement l''icône prévue au niveau du Bureau PGI.')
        else TvaModifEnc ;                  // Modification des bases HT
      end;
    7648 : TvaModifAcc;                     // Régularisation sur acomptes

    // Editions de TVA
    7650 : EditionTvaFac(True);     // Gestion directe des factures
    7651 : EditionTvaFac(False);    // Factures par étapes
    7652 : EditionTvaAcE;           // Etat des acomptes
    7654 : TvaValidEnc;             // Validation des éditions de TVA
    7647 : EditionTvaHT;            // Etat de contrôle
    7660 : CPLanceFiche_SuiviTVA(); // Suivi de la TVA par les soldes


    //SDA le 12/12/2007 1492 : LanceDeclarationTVA;
    1493 : LanceDeclarationTVA ('723'); //rélevé trimestriel des livraisons intracommunautaires
    1494 : LanceDeclarationTVA ('725'); // liste annuelle des clients assujettis
    //Fin SDA le 12/12/2007

    // Gestion du prorata de TVA
    7655 : CCLanceFiche_ParamTauxTVA;
    7656 : CCLanceFiche_EcrProrataTVA('First;999999103');
    7657 : CCLanceFiche_EcrProrataTVA('Last;999999105');
    7659 : CCLanceFiche_EcrTVAProratisees;
    140501 : if VH^.PaysLocalisation=CodeISOES then
                CP_LanceFicheESPCPQRMOD347; //Edition rapport pour modèle 347 (Chiffres d'affire superieure ou egal à 3K€ )
    140502 : if VH^.PaysLocalisation=CodeISOES then
                CP_LanceFicheESPCPQRJALTVA('TYPEETAT=ACHAT'); //journal de la TVA des achats
    140503 : if VH^.PaysLocalisation=CodeISOES then
                CP_LanceFicheESPCPQRJALTVA('TYPEETAT=VENTE') ; //journal de la TVA des achats //XMG 22/10/03
    140504 : if VH^.PaysLocalisation=CodeISOES then
                CP_LanceFicheESPCPQRTVA300;  //Modèles officiaux ESPAGNOLS TVA 320
    140505 : if VH^.PaysLocalisation=CodeISOES then
                CP_LanceFicheESPCPQRIRPF110; //Modèles officiaux ESPAGNOLS TVA 300
    140506 : if VH^.PaysLocalisation=CodeISOES then
                CP_LanceFicheESPCPQRIRPF115; //Modèles officiaux ESPAGNOLS TVA 390 //XVI 24/02/2005

//=====================
//==== Budget (12) ====
//=====================
    1530 : ParamRepartBUDGET('',False); {JP 21/11/05 : FQ 16521 : Uniformisation 2/3 - eAgl : Clés de répartition mensuelles}
    // Ecritures
    15211 : MultiCritereMvtBud(taCreat,'G');  // Ecritures / Saisie par comptes
    15213 : MultiCritereMvtBud(taCreat,'S');  // Ecritures / Saisie par sections
    15217 : MultiCritereMvtBud(taCreat,'GS'); // Ecritures / Saisie comptes sur une section
    15219 : MultiCritereMvtBud(taCreat,'SG'); // Ecritures / Saisie sections sur un compte
    15221 : MultiCritereMvtBud(taConsult,''); // Ecritures / Visualisation
    15230 : MultiCritereMvtBud(taModif,'');   // Ecritures / Modification
    15225 : VisuConsoBudget('G','');          // Ecritures / Vision consolidée / Comptes/sections
    15227 : VisuConsoBudget('S','');          // Ecritures / Vision consolidée / Sections/Comptes
    15240 : DetruitBudgets;                   // Ecritures / Suppression

    // Structures
    15115 : MulticritereBudgene(taModif);  // Comptes
    15135 : MulticritereBudsect(taModif);  // Sections
    15155 : MulticritereBudjal(taModif);   // Budget
    15119 : SuppressionCpteBudG;           // Suppression / Comptes
    15139 : SuppressionCpteBudS;           // Suppression / Sections
    15159 : SuppressionJournauxBud;        // Suppression / Budgets
    15125 : PlanBudget('', False);         // Impression / Comptes
    15145 : PlanBudSec('','', False);      // Impression / Sections
    15165 : PlanBudJal('','',False);       // Impression / Budgets
    15117 : MulticritereBudgene(taModifEnSerie); // Modifications en série / Comptes budgétaires
    15137 : MulticritereBudsect(taModifEnSerie); // Modifications en série / Sections budgétaires
    15157 : MulticritereBudjal(taModifEnSerie);  // Modifications en série / Budgets
    15118 : ModifSerieTableLibre(fbBudGen) ;     // MODIF PACK AVANCE
    15138 : ModifSerieTableLibre(fbBudSec1) ;    // MODIF PACK AVANCE

    // Editions
    15321 : BalBudteGen;               // Balances / Par comptes Lek 230106
    15322 : BalBudteSec;               // Balances / Par sections
    15323 : BalBudteGenSec;            // Balances / Par comptes et par sections
    15324 : BalBudteSecGen;            // Balances / Par sections et par comptes
    15331 : BalEcartBudParCpte;
    15332 : BalEcartBudParSection;
    15333 : BalEcartBudParCpteEtSection;
    15334 : BalEcartBudParSectionEtCpte;

    15282 : CPLanceFiche_CPBROUILLARDBUD('G');  {FP 28/12/2005} //BrouillardBud('N');        // Brouillard / Brouillard par comptes budgétaires
    15284 : CPLanceFiche_CPBROUILLARDBUD('S');  {FP 28/12/2005} //BrouillardBudAna('N') ;
    15350 : LanceEtatLibre('UBU');   // Etats libres
    15360 : LanceListeLibre('UBU');   // Etats libres

    // Traitements
    15250 : MultiCritereValBud(True);       // Validation
    15260 : MultiCritereValBud(False);      // Dévalidation
    15270 : RecopieBudgetSimple;            // Recopie de budget
    15271 : RecopieBudgetMultiple;          // Fusion
    15272 : RevisionBudgetaireRealiser;     // Révision
    15161 : OuvreFermeCpte(False);          // Fermeture
    15163 : OuvreFermeCpte(True);           // Réouverture

    // Reporting
    15410 : MultiCritereRubriqueV2(taModif,True); // Rubriques
    15430 : ControleRubrique('BUDGET');           // Contrôle du paramétrage
    15440 : TotalRubriquesV2(True);               // Contrôle des montants
    15421 : SuppressionRubriqueV2(True);          // Suppression
    15422 : CPLanceFiche_CPRESTANABUD; //FP 29/12/2005
    15425 : ImpRubrique('',False,True);           // Impression des rubriques

    // Paramètres
    1496 : YYLanceFiche_Souche('BUD','','ACTION=MODIFICATION;BUD'); // Compteurs
    1498 : ParamStructureBudget(PRien);                             // Catégories budgétaires
    15412 : ParamCroiseCompte;                                      // Gestion des croisements
    1501 : CodeGeneVersCodeBudg;                                    // Automates de création des comptes / à partir des comptes généraux
    1502 : CreationCodBudSurRupture(fbGene);                        // Automates de création des comptes / à partir d'un plan de ruptures
    1504 : CreationDesCodesBudget(fbBudGen);                        // Automates de création des comptes / à partir de valeurs de tables libres
    1522 : CreationCodBudSurRupture(fbAxe1);                        // Automates de création des sections / à partir d'un plan de ruptures
    1528 : GenerationSectionBudgetaire; {FP FQ15648}     // Automates de création des sections / à partir des sections analytiques
    1524 : CreationDesCodesBudget(fbBudSec1);                       // Automates de création des sections / à partir de tables libres
    1529 : GenSecBudCategorie;                                      // Automates de création des sections / à partir de catégories budgétaires
    15411 : CreationRubrique;                                       // Automates de création des rubriques budgétaires
    15126 : ControleBudget(fbBudGen);                               // Contrôles / des comptes
    15146 : ControleBudget(fbAxe1);                                 // Contrôles / des sections
    15340 : MultiBudgets; // Lek 270606

    17141 : CpLanceFiche_TypeVisaMul;
    17142 : CpLanceFiche_CircuitMul;
    17143 : CpLanceFiche_MailsBAP('', '', '');

//    17500 : CPLanceFiche_GestionEtatSynthese;

    11301 : CpLanceFiche_AffectationCircuit;
    11304 : CpLanceFiche_MulBap(tym_Modif);
    11308 : CpLanceFiche_MulBap(tym_Suivi);
    11312 : CpLanceFiche_MulBap(tym_Relance);
    11314 : CpLanceFiche_MulBap(tym_Suppression);
    11316 : CpLanceFiche_MulBap(tym_Purge);
    11318 : CpLanceFiche_MulBap(tym_Tous); {JP 24/03/07 : FQ 19949}
    11320 : CpLanceFiche_HistoriqueMail;

//===========================================
//==== Amortissement (20-PGE) et (2-PCL) ====
//===========================================
{$IFDEF AMORTISSEMENT}
    // Structures
    2111 : AMLanceFiche_ListeDesImmobilisations ( '' );
    2115 : AMLanceFiche_ListeDesImmobilisations ( '' , True); // Historique
    2150 : begin
           MAJDateFinContratImmos;           // FQ 19156
           CPLanceFiche_EtatsChaines('AM');
           end;
    // Editions
    2549 : AMLanceFiche_Edition ( 'IPM', 'IPM' );
    2550 : AMLanceFiche_Edition ( 'ILD', 'ILD' ); // mbo ajout dépréciation reprise 17.11.05
    2551 : AMLanceFiche_Edition ( 'ILC', 'ILC' ); // Tga ajout liste des changements 07.12.05
    2552 : AMLanceFiche_Edition ( 'ILV', 'ILV' ); // Pgr - FQ 15637 - Ajout liste des valorisations 27.12.05
    2542 : AMLanceFiche_Edition ( 'IAC', 'IAC' );
    2541 : AMLanceFiche_Edition ( 'ILS', 'ILS' ); // PCL
    20276: AMLanceFiche_Edition ( 'IET', 'IET' ); // TGA 06/02/2006 Etiquettes
//    2546 : AMLanceFiche_Edition ( 'AMN', 'AMN' );
    2543 : AMLanceFiche_Edition ( 'ISO', 'ISO' );
    2544 : AMLanceFiche_Edition ( 'ITE', 'ITE' );
    20222 : AMLanceFiche_Edition ( 'IMU', 'IMU' );
    20232 : AMLanceFiche_Edition ( 'ITN', 'ITN' );
    20233 : AMLanceFiche_Edition ( 'ITS', 'ITS' ); // BTY 10/06 Suramortissement des primes d'équipement
    20234 : AMLanceFiche_Edition ( 'ITF', 'ITF' );
    20236 : AMLanceFiche_Edition ( 'ITR', 'ITR' );
    20238 : AMLanceFiche_Edition ( 'ITO', 'ITO' );
    20239 : AMLanceFiche_Edition ( 'IRE', 'IRE' );
    20240 : AMLanceFiche_Edition ( 'ITV', 'ITV' ); // Tga liste des amortissements variables 16/02/06
    20241 : AMLanceFiche_Edition ( 'ITB', 'ITB' ); // BTY 10/06 Amortissement des subventions
    20252 : begin
            MAJDateFinContratImmos;           // FQ 19156
            AMLanceFiche_Edition ( 'ITP', 'ITP' );
            end;
    2547 : AMLanceFiche_Edition ( 'IPR', 'PRE' );
    2548 : AMLanceFiche_Edition ( 'IPR', 'PRF' );
    20272 : AMLanceFiche_Edition ( 'IVE', 'IVE' );
    20274 : AMLanceFiche_Edition ('IVL','IVL' );
    2180 : AMLanceFiche_StatPrevisionnel;
    20262 : AMLanceFiche_Edition ('IEG','IEG' );
    // BTY 02/06 FQ 13035 Dotations simulées
    20266 : AMLanceFiche_Edition ('IDS','IDS' );
        // MBO 03/07 FQ 17512 réintégrations/Déductions diverses
    20281 : AMLanceFiche_Edition ('IRD','IRD' );
    // MBO 03/07 FQ 17512 quote-part personnelle
    20282 : AMLanceFiche_Edition ('ITQ','ITQ' );
    // MBO 04/07 Ajout Remplacement d'un composant
    20283 : AMLanceFiche_Edition ('IRC','IRC' );
    // Traitements
    2410 : AMLanceFiche_ListeDesIntegrations ('');
    2412 : AMLanceFiche_ListeImmobilisationsARD (taModif, ''); // BTY 07/07 Gestion des ARD
    2415 : EtatRapprochementCompta ;
    2416 : AfficheClotureImmo ;
    2417 :
      begin
        AMLanceFiche_AnnulationCloture;
        RetourForce:=true;
      end;
    2418 : UpdateBaseImmo ;    //ajout TD le 8/6/2001 menu Recalcul des plans

    // Paramètres
    //30/10/07 BTY Modif rubrique Agricole dans SCO_IMMOBILISATIONS 2340 : ParamSociete(False,'','SCO_IMMOBILISATIONS','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000);
    2340 : ParamSociete(False,'','SCO_COMPTES;SCO_FOUR2;SCO_INTEGRATION;SCO_JOURSORTIE','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000);
    2210 : AMLanceFiche_ComptesAssocies ( taModif );
    2342 : ParamTable('TILIEUGEO',taCreat,2342000,PRien) ;
    2345 : ParamTable('TIMOTIFCESSION',taCreat,2345000,PRien) ;
    2347 : SaisieDesCoefficientsDegressifs ;
    //TGA 24/03/2006 2348 : ParamTable('AMREGROUPEMENT',taCreat,0,PRien);
    2348 : AMLanceFiche_SaisieRegroupement('IGI');
    // Utilitaires
    2370 : LanceVerificationAmortissement;
    2375 : LanceVerificationAmortissement;

    20511 : SavImo2PGI ('',False );
    20512 : AMLanceFiche_AmExport; // ExporteAmortissement ;
    20513 : AMLanceFiche_AmExport ('RI'); // Export enrichi BTY 05/07    
{$ENDIF}

//======================
//==== Analyses (8) ====
//======================
    // Cubes
    8305  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','') ;
    8355  : AGLLanceFiche('CP','CPECRITURE_OLAP','','','') ; // Cube OLAP
    8307  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','CONTREPARTIE') ;
    8357  : AGLLanceFiche('CP','CPANALYTIQ_OLAP','','','') ; // Cube OLAP
    8310  : AGLLanceFiche('CP','CPANALYTIQ_CUBE','','','') ;
    8312  : CPLanceFiche_BalAgeeCube('','','') ;
    8315  : CPLanceFiche_BalAgeeCube('','','IAS14') ;
    // Analyses statistiques
    8405  : AGLLanceFiche('CP','CPECRITURE_TOBV','','','') ;
    8410  : AGLLanceFiche('CP','CPANALYTIQ_TOBV','','','') ;
    8412  : AGLLanceFiche('CP','CPANAGENE_TOBV','','','') ;
    // Analytiques multi-axes
    8205  : AGLLanceFiche('CP','CPGENERECROISEAXE','','','') ;
    8210  : AGLLanceFiche('CP','CPCROISEAXE_CUBE','','','') ;
    8215  : AGLLanceFiche('CP','CPCROISEAXE_STAT','','','') ;
    8220  : AGLLanceFiche('CP','CPCROISEAXE_ETAT','','','') ;

// BPY le 21/06/2004 => Restitutions multi sociétés
    // MULTIDOSSIER
    // Cubes
    8511  : AGLLanceFiche('CP','CPECRITURE_CUBEM','','','');
    8512  : AGLLanceFiche('CP','CPECRITURE_CUBEM','','','CONTREPARTIE') ;
    8513  : AGLLanceFiche('CP','CPANALYTIQ_CUBEM','','','') ;
    8514  : CPLanceFiche_BalAgeeCubeM('','','') ;
    8515  : CPLanceFiche_BalAgeeCubeM('','','IAS14') ;
    // Analyses statistiques
    8521  : AGLLanceFiche('CP','CPECRITURE_TOBVM','','','');
    8522  : AGLLanceFiche('CP','CPANALYTIQ_TOBVM','','','') ;
    8523  : AGLLanceFiche('CP','CPANAGENE_TOBVM','','','') ;
    // Analytiques multi-axes
    // non utilisé pour le moment !!
// Fin BPY

//=======================================
//==== Structures et paramètres (17) ====
//=======================================
    // Structures
    7112 : CCLanceFiche_MULGeneraux('C;7112000'); //MulticritereCpteGene(taModif);
    7145 : CPLanceFiche_MULTiers('C;7145000');    //MultiCritereTiers(taModif);
    7178 : CPLanceFiche_MULSection('C;7178000');  //MulticritereSection(taModif);
    7211 : CPLanceFiche_MULJournal('C;7211000');  //MulticritereJournal(taModif);

    7117 : CPLanceFiche_ChgNatureGene;
    7118 : CCLanceFiche_MULGeneraux('S;7118000');  // Suppressions // Comptes généraux
    7151 : CPLanceFiche_MULTiers('S;7151000');     // Suppressions // Comptes auxiliaires
    7184 : CPLanceFiche_MULSection('S;7184000');   // Suppressions // Sections analytiques
    7214 : CPLanceFiche_MULJournal('S;7214000');   // Suppressions // Journaux

    7124 : CCLanceFiche_MULGeneraux('F;7124000');  // Ouverture/Fermeture // Fermeture des généraux
    7127 : CCLanceFiche_MULGeneraux('O;7127000');  // Ouverture/Fermeture // Ouverture des généraux
    7157 : CPLanceFiche_MULTiers('F;7157000');     // Ouverture/Fermeture // Fermeture des auxiliaires
    7160 : CPLanceFiche_MULTiers('O;7160000');     // Ouverture/Fermeture // Ouverture des auxiliaires
    7190 : CPLanceFiche_MULSection('F;7190000');   // Ouverture/Fermeture // Fermeture des sections
    7193 : CPLanceFiche_MULSection('O;7193000');   // Ouverture/Fermeture // Ouverture des sections
    7220 : CPLanceFiche_MULJournal('F;7220000');   // Ouverture/Fermeture // Fermeture des journaux
    7223 : CPLanceFiche_MULJournal('O;7223000');   // Ouverture/Fermeture // Ouverture des journaux
    3150 : CCLanceFiche_LongueurCompte;

    7133 : CPLanceFiche_PLANGENE('');     // Impression / Comptes généraux
    7166 : CPLanceFiche_PLANAUXI('');     // Impression / Comptes auxiliaires
    7199 : CPLanceFiche_PLANSECT('', ''); // Impression / Sections analytiques
    7229 : CPLanceFiche_PLANJOUR('');     // Impression / Journaux
    //YMO 23/03/2006 Menus du specif CERIC
    7240 : CPLAnceFiche_AFFEnCours('', '', '') ;       // Affaire en cours

    // Modifications en série
    // L'aide en ligne est modifiée pour la version 6 corrective.
    // Il faudra remettre les numéros d'aide lorsque le fichier d'aide contiendra ces numéros
    7115 : CCLanceFiche_MULGeneraux('SERIE;150'); // 7115000'); // MulticritereCpteGene(taModifEnSerie);
    7148 : CPLanceFiche_MULTiers('SERIE;150'); // 7148000');    // MultiCritereTiers(taModifEnSerie);
    7181 : CPLanceFiche_MULSection('SERIE;150'); // 7181000');  // MulticritereSection(taModifEnSerie);

    7116 : ModifSerieTableLibre(fbGene) ;       // MODIF PACK AVANCE
    7149 : ModifSerieTableLibre(fbAux) ;        // MODIF PACK AVANCE
    7182 : ModifSerieTableLibre(fbSect) ;       // MODIF PACK AVANCE

    // Société
    1193 : ParamTable('ttNivCredit',taCreat,1190040,PRien) ;

    1104 : begin
             ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
             RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
           end;

{$IFDEF EAGLCLIENT}
{$ELSE}
    1197 : CopiTableLibre ;                               // MODIF PACK AVANCE
    1526 : CopiePlanSousSectionToTableLibre ;             // MODIF PACK AVANCE
{$ENDIF}
    1155 : ParamTable('TTLANGUE',taCreat,1155000,PRien) ; // MODIF PACK AVANCE

    1195 : AGLModifLibelleChampLibre('','') ; // FQ 16256 SBO 01/08/2005
    1105 : FicheAxeAna ('') ;
    1110 : FicheEtablissement_AGL(taModif) ;
    1120 : ParamTable('ttFormeJuridique',taCreat,1120000,PRien) ;
    1122 : ParamTable('ttCivilite',taCreat,1122000,PRien) ;
    1125 : ParamTable('ttFonction',taCreat,1125000,PRien) ;
    1135 : OuvrePays ;
    1140 : FicheRegion('','',False) ;
    1145 : OuvrePostal(PRien) ;
    1150 : FicheDevise('',tamodif,False) ;
    1165 : ParamTable('TTREGIMETVA',taCreat,1165000,PRien) ;
    1170 : ParamTvaTpf(true) ;   // Tables // TVA par régime fiscal
    1175 : ParamTvaTpf(false) ;  // Tables // TPF par régime fiscal
    1185 : FicheModePaie_AGL('');
    1190 : FicheRegle_AGL('',FALSE,taModif) ;
    1191 : ParamTable('ttQualUnitMesure',taCreat,1190030,PRien) ;

    1194 : CPLanceFiche_ParamTablesLibres;  // Tables libres // Personalisation
    1196 : CPLanceFiche_ChoixTableLibre;    // Tables libres // Saisie
    1198 : SynchroTaLi;                     // Tables libres // Synchronisation

    // Paramètres
    1290 : YYLanceFiche_Exercice('0');  // Exercice
{$IFDEF ADRESSE}
    1396 : YY_LanceParamAdresse;    // Gestion des adresses.
    1397 : YY_LanceSaisieAdresse;
{$ENDIF}
    // Compteurs
    1300 : YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT'); //Compteurs
    1301 : YYLanceFiche_Souche('REL','','ACTION=MODIFICATION;REL'); //Compteurs

    //Multidossier
    1750 : ParamRegroupementMultiDossier ;

    {$IFNDEF EAGLCLIENT}
    1755 : YYLanceFiche_Bundle( 'YY', 'YYBUNDLE', '', '', '') ;
    {$ENDIF EAGLCLIENT}
    1760 : CPLanceFiche_LiensEtab('','','') ;
    1765 : CPLanceFiche_LiensSociete('','','') ;

    {$IFDEF TAXES}
    1770 : YYLanceFiche_ModeleTaxe; //SDA le 12/02/2007
    1775 : YYLanceFiche_TypesTaxes; //SDA le 12/02/2007
    1780 : YYLanceFiche_ModeleCatRegime; //SDA le 12/02/2007
    1790 : YYLanceFiche_ModeleCatType; //SDA le 15/02/2007
    1795 : YYLanceFiche_CompteParCatTaxe; //SDA le 02/03/2007
    1800 : YYLanceFiche_Detailtaxes;//SDA le 26/02/2007
    {$ENDIF}

  {JP 30/09/05 : Cette unité est compatible eAGL}
    1325 : CCLanceFiche_Correspondance('GE');	// Plan de correspondance Généraux
    1330 : CCLanceFiche_Correspondance('AU');	// Plan de correspondance Auxiliaire
    1345 : CCLanceFiche_Correspondance('A1');	// Plan de correspondance Axe analytique 1
    1350 : CCLanceFiche_Correspondance('A2');	// Plan de correspondance Axe analytique 2
    1355 : CCLanceFiche_Correspondance('A3');	// Plan de correspondance Axe analytique 3
    1360 : CCLanceFiche_Correspondance('A4');	// Plan de correspondance Axe analytique 4
    1365 : CCLanceFiche_Correspondance('A5');	// Plan de correspondance Axe analytique 5
    {JP 30/09/05 : Semble inutile car jamais utilisé (vu avec OG et EC
     1426 : CCLanceFiche_Correspondance('ZG') ; }

{$IFDEF EAGLCLIENT}
{$ELSE}
    1315 : FichePlanRef(1,'',taModif) ;

    1375 : PlanRupture('RUG') ;
    1380 : PlanRupture('RUT') ;
    1395 : PlanRupture('RU1') ;
    1400 : PlanRupture('RU2') ;
    1405 : PlanRupture('RU3') ;
    1410 : PlanRupture('RU4') ;
    1415 : PlanRupture('RU5') ;
    1430 : CPLanceFiche_MulGuide('','',''); //ParamGuide('','NOR',taModif) ;
    1435 : ParamScenario('','') ;

    1485 : ParamEtapeReg(True,'',True) ;
    1490 : ParamEtapeReg(False,'',True) ;
    1491 : CircuitDec ;
{$ENDIF}
    1425 : ParamLibelle ;          // Libellés automatiques
    1450 : ParamPlanAnal('');      // Analytiques / Structures analytiques
    1455 : ParamRepartCle;         {JP 08/06/05 : Répartition des clefs analytiques}
    1460 : ParamVentilType(PRien); // Analytiques / Ventilations types
    1465 : CPLanceFiche_CPMODRESTANA; //FP 29/12/2005
    1470 : CPLanceFiche_CPRESTANAECR;//FP 29/12/2005
    7574 : CCLanceFiche_ParamRelance('RTR');
    7577 : CCLanceFiche_ParamRelance('RRG');

    // Ecritures
    7352 : CPLanceFiche_MulGuide('','','');   // Guide
    7355 : CPLanceFiche_Scenario('','');      // Scénario
    1701 : LanceFicheIdentificationBancaire;
    17420 : CPLanceFiche_MasqueMul('','','ACTION=MODIFICATION;CMS_TYPE=SAI') ;

    // Banques
    {JP 10/03/04 : Ajout d'un ParamSoc qui détermine si l'on utilise la Tréso ...}
      1705 : if EstComptaTreso then TRLanceFiche_Banques('TR','TRCTBANQUE', '', '', '')
                               else FicheBanque_AGL('',taModif,0);
      1706 : TRLanceFiche_Agence('TR','TRAGENCE', '', '', '');
      1707 : TRLanceFiche_BanqueCP('TR','TRMULCOMPTEBQ','','','');
    1720 : ParamTable('ttCFONB',tacreat,500005,PRien);        // Codes CFONB
    1725 : ParamCodeAFB;                                 // Code AFB
{$IFDEF EAGLCLIENT}

{$ELSE}
    1730 : ParamTable('ttCodeResident',taCreat,500007,PRien) ;
    1732 : ParamTeletrans ;
{$ENDIF}
// GP le 18/08/2008 : 23236
    1736 : TRLanceFiche_MulCIB('TR', 'TRMULCIB', '','', '');
    1738 : TRLanceFiche_ParamRegleAccro('TR', 'TRFICHEACCROCHAGE', '', '', '');
//    1739 : TRLanceFiche_CIB('TR','TRCIB','','','');
    1739 : TRLanceFiche_CIB('TR','TRCIB','','','ACTION=MODIFICATION;' + tc_CIB);      // Code InterBancaire

    // Documents
{$IFDEF EAGLCLIENT}

{$ELSE}
    1235 : EditDocument('L','REL','',TRUE) ;
{ FQ 19667 BVE 09.05.07
    1240 : EditDocument('L','RLV','RLV',FALSE) ; }
    1240 : EditDocument('L','RLV','RLW',FALSE) ;
{ END FQ 19667 }
    1245 : EditDocument('L','RLC','',True) ;
{$ENDIF}
    1247 : EditEtat('E','SAI','',True,nil,'','') ;
    1248 : EditEtat('E','BAP','',True,nil,'','') ;

   {JP 14/01/08 : Reprise de ce qui était fait dans MDispMp.Pas}
   {$IFDEF EAGLCLIENT}
    1266 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLP', '',True, nil, '', 'Lettres prélèvements');
    1267 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLC', '',True, nil, '', 'Lettres chèques');
    1268 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLT', '',True, nil, '', 'Traites LCR / BOR');
    1269 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLV', '',True, nil, '', 'Lettres virements');
    {$ELSE}
    1266 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLP', '',True, nil, '', 'Lettres prélèvements')
                                                               else EditDocument('L','LPR','',True) ;
    {JP 24/01/06 : FQ 17234 : Générateur de documents ou d'états en fonction du ParamSoc}
    1267 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLC', '',True, nil, '', 'Lettres chèques')
                                                               else EditDocument('L','LCH','',True) ;
    1268 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLT', '',True, nil, '', 'Traites LCR / BOR')
                                                               else EditDocument('L','LTR','',True) ;
    {JP 04/06/07 : FQ 19416 : Générateur de documents ou d'états en fonction du ParamSoc}
    1269 : if GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then EditEtat('E', 'CLV', '',True, nil, '', 'Lettres virements')
                                                               else EditDocument('L','LVI','',True) ;
    {$ENDIF}
    1257 : EditEtat('E','BOR','',True,nil,'','') ;

    {$IFNDEF TT}
    // Editions paramétrable
    1280 : EditEtatS5S7('E','JDP','',TRUE,nil,'','') ; // Journal divisionnaire // Modif SBO changement de nature JDI --> JDP
    1281 : EditEtatS5S7('E','JAC','',TRUE,nil,'','') ; // Journal centralisateur
    1282 : EditEtatS5S7('E','JPE','',TRUE,nil,'','') ; // Journal périodique
    1283 : EditEtatS5S7('E','JAL','',TRUE,nil,'','') ; // Journal général
    1284 : EditEtatS5S7('E','GLP','',TRUE,nil,'','') ; // Grand livre général
    1285 : EditEtatS5S7('E','GLB','',TRUE,nil,'','') ; // Grand livre auxiliaire // Modif SBO changement de nature GLA --> GLB
    1286 : EditEtatS5S7('E','BGE','',TRUE,nil,'','') ; // Balance générale
    1287 : EditEtatS5S7('E','BAU','',TRUE,nil,'','') ; // Balance auxiliaire
    1288 : EditEtatS5S7('E','BVE','',TRUE,nil,'','') ; // Balance ventilée
    1289 : EditEtatS5S7('E','JTA','',TRUE,nil,'','') ; // Tableau d'avancement
    1291 : EditEtatS5S7('E','RES','',TRUE,nil,'','') ; // Document de synthèse - Résultat
    1292 : EditEtatS5S7('E','BIL','',TRUE,nil,'','') ; // Document de synthèse - Actif
    1293 : EditEtatS5S7('E','SIG','',TRUE,nil,'','') ; // Document de synthèse - SIG
    1270 : EditEtatS5S7('E','UCO','',True,nil,'','') ;
    1271 : EditEtatS5S7('E','UBU','',True,nil,'','') ;
    {$ENDIF}
//=======================================
//==== Administrations / Outils (18) ====
//=======================================
    // Installation
{$IFDEF EAGLCLIENT}

{$ELSE}
    // Perso/Traduc
    1205 : ParamTraduc(TRUE,Nil) ;   // MODIF PACK AVANCE
    1210 : ParamTraduc(FALSE,NIL) ;  // MODIF PACK AVANCE
{$ENDIF}

    // Société
    3155 :
      begin
        RAZSociete ;
        RetourForce := True;
      end;
    3247 : CPLanceFiche_EnregStd;
{$IFDEF EAGLCLIENT}
    3140 : if ctxPCL in V_PGI.PGIContexte then LanceAssistantComptaPCL(False);
{$ELSE}
    3130 : GestionSociete(PRien,@InitSociete,nil) ;
    3140 : if ctxPCL in V_PGI.PGIContexte then LanceAssistantComptaPCL(False) else LanceAssistSO ;
    { YMO 29/12/2005 FQ 10471 Suppression du menu Recopie société à société
    3145 : RecopieSoc/(PRien,True) ; }
    // PCL
    3185 : VisuLog1 ; // FQ17895 YMO 27/06/06 Remise en place menu 'Suivi d'activité' en CWAS ** Suppression eagl 08/12/2006
{$ENDIF}
    3172 : YYLanceFiche_ProfilUser;
    3248 : CPLanceFiche_AlignementSurStandard;
    // Utilisateurs et accès
    3165  : FicheUserGrp ;
    3170 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;
//GP le 30/08/2008 : donner l'accès au etats 2/3 en 2/3 (pb ancien état de tva jamais accessibles
//60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','8;9;11;12;13;14;16;17;18;20;26;27;324;361');
{$IFDEF EAGLCLIENT}
    60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','8;9;11;12;13;14;16;17;18;20;26;27;324;361');
{$ELSE}
    60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','8;9;11;12;13;14;16;17;18;20;26;27;324;361;363');
{$ENDIF}


    3180 : ReseauUtilisateurs(False) ;

    3195 : ReseauUtilisateurs(True) ;
{$IFDEF CERTIFNF}
    { BVE 03.09.07 : Journal des évenements }
    3169 : CPLanceFicheCPJNALEVENT('','','');
{$ENDIF}

    {JP 16/02/04 : pas de tof associée, les traitements sont dans le script}
    3198 : AGLLanceFiche('CP', 'CPREINITFNCT', '', '', '');
    // Outils
    7120 : CPLanceFiche_Reclassement ; // Outil de reclassement des données (PCL)
    3133 :
      begin
      if ctxPCL in V_PGI.PGIContexte then
        begin
          stAExclureFavoris := '';
          TripotageMenuModePCL ( [52,53,56], stAExclureFavoris );
          TripotageMenuAmortissement (stAExclureFavoris );
          ParamFavoris('54;55;96', stAExclureFavoris, False, True);
        end
        else
        begin
          stAExclureFavoris := '';
          TripotageMenuModePGE ( [9,11,13,16,14], stAExclureFavoris );
          TripotageMenuAmortissement (stAExclureFavoris );
          ParamFavoris('', stAExclureFavoris, False, True);
        end ;
      end;
    3205 : CCLanceFiche_LongueurCompte ;
    7119 : CRazModePointage ;          // Changement du mode de pointage
    54150 : TraitementLibelle; // ajout me 10/11/2005

    3217 : Commun.CreerAgence; {JP 10/03/04 : mise ne place de l'initialisation des agences}
    3212 : CPLanceFiche_CPANOECR ;
{$IFDEF EAGLCLIENT}
    {$IFDEF CONTREPARTIE}
    3215 : CCLanceFiche_ReparationFic('');   // ajout me
    {$ENDIF}
{$ELSE}
    3210 : ControlFic ;
    3215 : ReparationFic ;
    3220 : ReindexSoc ;
    3221 : CPLanceFiche_RubImport;

    56180 : CCLanceFiche_ModifEntPie; // Modification d'entête de pièces (PCL)
    56190 : CPLanceFiche_CPPurgeExercice;
{$IFDEF CERTIFNF}
    { BVE 03.09.07 : Journal des évènements }
    56205 : CPLanceFicheCPJNALEVENT('','','');
{$ENDIF}
{$ENDIF}

    // Comptable
    3227 : CPLanceGestionAna;
    3230 : CPLanceFiche_MajTVA(''); {09/11/2007 YMO Outil Maj TVA}
    3245 : AnnuleclotureComptable(TRUE);

//=======================
//=== Menu Pop Compta ===
//=======================
    25710 : CCLanceFiche_MULGeneraux('P;7112000');  // Comptes généraux
    25720 : CPLanceFiche_MULTiers('P;7145000');     // Comptes auxiliaires
    25730 : CPLanceFiche_MULSection('P;7178000');   // Sections analytiques
    25740 : CPLanceFiche_MULJournal('P;7205000');   // Journaux
    25750 : if CtxPcl in V_Pgi.PgiContexte then
              CPLanceFiche_CPRechercheEcr(False)
            else
              MultiCritereMvt(taConsult,'N',False); // Consultation des écritures

    25755 : OperationsSurComptes('','','') ;  	    // Consultation des comptes
    25760 : ResultatDeLexo ;
    25770 : begin
              ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
              RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
            end;

//=======================
//=== Menu Pop Budget ===
//=======================
    26160 : MulticritereBudgene(taConsult);
    26165 : MulticritereBudsect(taConsult) ;
    26170 : MulticritereBudjal(taConsult) ;
    26175 : MultiCritereMvtBud(taConsult,'');

    26180 : CCLanceFiche_MULGeneraux('C;7112000');  // Comptes généraux
    26185 : CPLanceFiche_MULSection('C;7178000');   // Sections analytiques
    26190: ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
      RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);

//=============================
//==== Suivi clients (37)  ====
//=============================
    370101 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=MPF') ; //Mise en portefeuille
    370102 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=REC') ; //Remise à e'encaissement
    370103 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=REE') ; //Remise à l'escompte
    370104 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=COC') ; //Confirmation rem8ise encaissement
    370105 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=COE') ; //Confirmation remise escompte
    370106 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=IMC') ; //Impayé remise encaissement
    370107 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=IME') ; //Impayé remise escompte
    370108 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=PIM') ; //règlement impayé
    370109 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=RNG') ; //Re-negotation
    370110 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=ENC') ; //Encaissement directe
    370111 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=RJC') ; //Reject remise encaissement
    370112 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=RJE') ; //Reject remise escompte
    370113 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=ENC;OPERATION=FUS') ; //Fusion traites. //XVI 24/02/2005
    37105 : SuivAcceptation(prClient);     // Suivi acceptation
    37108 : SelectSuiviMP(smpEncTraEdt) ; // Préparation lettres traites client
    37110 : GenereSuiviMP(smpEncTraEdtNC);// Génération Lettres-traite clients
    37115 : GenereSuiviMP(smpEncTraEdt) ; // Génération Lettres-traite clients
    37120 : SaisieEffet(OnEffet) ;
    37124 : SelectSuiviMP(smpEncTraPor) ; // Préparation mise en portefeuille client
    37125 : GenereSuiviMP(smpEncTraPor) ; // Mise en portefeuille traite acceptée
    37129 : SelectSuiviMP(smpEncTraEnc) ; // Préparation encaissement traite client
    37130 : GenereSuiviMP(smpEncTraEnc) ; // Encaissement traites clients Voir le smpEncTraBqe
    37134 : SelectSuiviMP(smpEncTraEsc) ; // Préparation escompte traite client
    37135 : GenereSuiviMP(smpEncTraEsc) ; // Eescompte traites clients
    37137 : SelectSuiviMP(smpEncTraBqe) ; // Preparation remise en banque traites
    37138 : GenereSuiviMP(smpEncTraBqe) ; // Remise en banque traites clients
    37140 : ExportCFONBMP(smpEncTraEsc) ; // Export cfong traite escompte
    37145 : ExportCFONBMP(smpEncTraEnc) ; // Export cfong traite encaissement

    // Prélévements
    37203 : SelectSuiviMP(smpEncPreEdt) ;   // Préparation lettre prélèvements
    37205 : GenereSuiviMP(smpEncPreEdtNC);  // Edition des lettres-prélèvement
    37204 : GenereSuiviMP(smpEncPreEdt) ;   // Edition et comptabilisation des lettres-prélèvement
    37207 : SelectSuiviMP(smpEncPreBqe) ;   // Comptabilisation prélèvement // Préparation
    37208 : GenereSuiviMP(smpEncPreBqe) ;   // Comptabilisation prélèvement // Comptabilisation
    37210 : ;                               // Emissions de bordereaux clients
    37215 : ExportCFONBMP(smpEncPreBqe) ;   // Export CFONB

    // Chèques
    37250 : SaisieEffet(OnChq) ;          // Saisie des chèques en portefeuille
    37251 : SelectSuiviMP(smpEncChqPor) ; // Mise en portefeuille // Préparation
    37252 : GenereSuiviMP(smpEncChqPor) ; // Mise en portefeuille // Remise
    37253 : SelectSuiviMP(smpEncChqBqe) ; // Remise à l'encaissement // Préparation
    37254 : GenereSuiviMP(smpEncChqBqe) ; // Remise à l'encaissement // Remise
    37256 : ;                             // Emission de bordereaux

    // Cartes bleues
    37260 : SaisieEffet(OnCB) ;          // Saisie de cartes bleues en portefeuille
    37261 : SelectSuiviMP(smpEncCBPor) ; // Mise en portefeuille // Préparation
    37262 : GenereSuiviMP(smpEncCBPor) ; // Mise en portefeuille // Remise
    37263 : SelectSuiviMP(smpEncCBBqe) ; // Remise à l'encaissement // Préparation
    37264 : GenereSuiviMP(smpEncCBBqe) ; // Remise à l'encaissement // Remise
    37266 : ;                            // Emission de bordereaux

    // Gestion par lots
{$IFDEF EAGLCLIENT}
    // A FAIRE
{$ELSE}
    37305 : AffecteBanquePrevi(True,taCreat) ;  // Création d'un lot
    37310 : AffecteBanquePrevi(True,taModif) ;  // Modification d'un lot
    37312 : UpdateLibelleToutModele ;           // Suppression d'un lot
{$ENDIF}
//    37305 : AffecteBanquePrevi(True,taCreat) ;  // Création d'un lot
//    37310 : AffecteBanquePrevi(True,taModif) ;  // Modification d'un lot
//    37312 : UpdateLibelleToutModele ;           //Suppression lot
    37315 : GenereSuiviMP(smpEncTous)   ;       // Encaissements

    // Relances clients
    37405 : RelanceClient(True,True) ;    // relance manuelle sur traites
    37410 : RelanceClient(True,False) ;   // relance manuelle sur autres modes paiement
    37415 : RelanceClient(False,True) ;   // relance auto sur traites
    37420 : RelanceClient(False,False) ;  // relance auto sur autres modes paiement
    // Editions
    37515 : AGLLanceFiche('CP', 'EPECHEANCIER','','ECH','ECH'); //FQ17255 YMO 06/01/2005
    37520 : CPLanceFiche_JustSold(smpEncTous);// JustSolde(smpEncTous) ; Lek 190106
    37525 : CPLanceFiche_BalanceAgee;
    37530 : CPLanceFiche_BalanceVentilee;  // Balance ventilée
    37531 : AGLLanceFiche('CP', 'SUIVICLIENT', '', '', ''); // Statistiques

    // Autres traitemens
    37501 : SelectSuiviMP(smpEncDiv);                         // Préparation Encaissements divers // SBO 20/04/007 : ajout du menu préparation    
    37502 : GenereSuiviMP(smpEncDiv);                         // Encaissements divers
    37505 : ExportCFONBBatch(False,True,tslAucun);            // Emissions de bordereaux banque
    37510 : ModifEcheMP;                                      // Modifications des échéances
    37535 : AGLLanceFiche('CP', 'PROFILUSER', '', '', 'CLI'); // Profils utilisateur
    37545 : ExportCFONBBatch(True,True,tslAucun,smpEncTous);  // Export CFONB
    37550 : SaisieEffet(OnBqe,srCli);                         // Saisie de trésorerie - SUPPRIME
    37555 : ReleveFacture;                                    // Relevés de factures - SUPPRIME

    37560 : CPLanceFiche_MulEncaDeca ( '', '', 'FLUX=ENC' ) ;                       // Génération des enca/deca
    37565 : CPLanceFiche_MulParamGener ( '', '', 'ACTION=MODIFICATION' ) ;  // Paramétrage des scénarios
    37570 : CP_LanceFicheExport('CLI;');{JP 02/05/05 : FQ 15333}
    37580 : SelectSuiviMP(smpCompenCli);         //FP 21/02/2006
    37585 : GenererCompensation(smpCompenCli);   //FP 21/02/2006
    37590 : CPLanceFiche_GenereEncaDeca( 'FLUX=ENC' ) ;

//==================================
//==== Suivi fournisseurs (38)  ====
//==================================
    380101 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=DEC;OPERATION=MPF') ; //Mise en portefeuille
    380102 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=DEC;OPERATION=DEC') ; //Decaissement directe
    380103 : if VH^.PaysLocalisation=CodeISOES then
                CPLanceFiche_MulEncaDeca('','','ACTION=MODIFICATION;Flux=DEC;OPERATION=FUS') ; //Fusion des traites //XVI 24/02/2005
    // B.O.R.
    38104 : SelectSuiviMP(smpDecBorEdt);   // Lettres-BOR // Préparation
    38105 : GenereSuiviMP(smpDecBorEdtNC); // Lettres-BOR // Edition des lettres BOR
    38110 : GenereSuiviMP(smpDecBorEdt);   // Lettres-BOR // Edition et comptabilisation des lettres-BOR
    38114 : SelectSuiviMP(smpDecTraPor);   // Mise en portefeuille // Préparation
    38115 : GenereSuiviMP(smpDecTraPor);   // Mise en portefeuille // Mise en portefeuille
    38119 : SelectSuiviMP(smpDecBorDec);   // Remise en banque // Préparation
    38120 : GenereSuiviMP(smpDecBorDec);   // Remise en banque // Remise
    38125 : ExportCFONBMP(smpDecBorDec);   // Export CFONB

    // Virements
    38202 : SelectSuiviMP(smpDecVirEdt) ;   // Lettres-virement // Préparation
    38206 : GenereSuiviMP(smpDecVirEdtNC) ; // Lettres-virement // Edition des lettres-virement
    38203 : GenereSuiviMP(smpDecVirEdt) ;   // Lettres-virement // Edition et comptabilisation des lettres-virement
    38204 : SelectSuiviMP(smpDecVirBqe) ;   // Comptabilisation virement // Préparation
    38205 : GenereSuiviMP(smpDecVirBqe) ;   // Comptabilisation virement // Comptabilisation
//    38210 :;                                // Emissions de bordereaux // SUPPRIME
    38215 : ExportCFONBMP(smpDecVirBqe) ;   // Export CFONB

    // Virements internationaux
    38282 : SelectSuiviMP(smpDecVirInEdt) ;   // Lettres-virement internationale // Préparation
    38286 : GenereSuiviMP(smpDecVirInEdtNC) ; // Lettres-virement internationale // Edition des lettres-virement internationale
    38283 : GenereSuiviMP(smpDecVirInEdt) ;   // Lettres-virement internationale // Edition et comptabilisation des lettres-virement internationale
    38284 : SelectSuiviMP(smpDecVirInBqe) ;   // Comptabilisation virement internationale // Préparation
    38285 : GenereSuiviMP(smpDecVirInBqe) ;   // Comptabilisation virement internationale // Comptabilisation
//    38290 : ;                                 // Emissions de bordereaux // SUPPRIME
    38297 : ExportCFONBMP(smpDecVirInBqe) ;   // Export CFONB

    // Gestion par lots
{$IFDEF EAGLCLIENT}
// A FAIRE
{$ELSE}
    38305 : AffecteBanquePrevi(False,taCreat) ;  // Création d'un lot
    38310 : AffecteBanquePrevi(False,taModif) ;  // Modification d'un lot
    38311 : SuivBAP ; // Bons à payer
    38312 : ; //Suppression lot - SUPPRIME
{$ENDIF}
    38315 : GenereSuiviMP(smpDecTous);          // Décaissements

    // Chèques
    38408 : SelectSuiviMP(smpDecChqEdt);   // Lettres-chèques // Préparation lettre chèque
    38409 : GenereSuiviMP(smpDecChqEdtNC); // Lettres-chèques // Edition des lettres-chèque
    38410 : GenereSuiviMP(smpDecChqEdt);   // Lettres-chèques // Edition et comptabilisation des lettres-chèque

    // Lettrage
    7509 : LanceLettrageMP(prFournisseur);             // Lettrage manuel
    7512 : LanceDeLettrageMP(prFournisseur);           // Délettrage
    7515 : RapprochementAutoMP('','',prFournisseur);   // Lettrage automatique
    7521 : RegulLettrageMP(True,False,prFournisseur);  // Régularisation de lettrage
    7525 : RegulLettrageMP(False,False,prFournisseur); // Différence de change
    7526 : RegulLettrageMP(False,True,prFournisseur);  // Ecart de conversion

    // Editions

    38515 : AGLLanceFiche('CP', 'EPECHEANCIER','','ECH','ECH'); // Echéancier  //FQ17255 YMO 06/01/2005
    38520 : CPLanceFiche_JustSold(smpDecTous);// JustSolde(smpDecTous); Lek 190106                       // Justificatif de solde
    38525 : CPLanceFiche_BalanceAgee;                     // Balance âgée
    38530 : CPLanceFiche_BalanceVentilee;                 // Balance ventilée
    38531 : AGLLanceFiche('CP', 'SUIVIFOU', '', '', '');  // Statistiques

    // Autres traitements
    38501 : SelectSuiviMP(smpDecDiv);                              // Préparation Décaissements divers // SBO 20/04/007 : ajout du menu préparation
    38502 : GenereSuiviMP(smpDecDiv);                              // Décaissements divers
    38505 : ExportCFONBBatch(False,FALSE,tslAucun);                // Emission borderau banque
    38510 : ModifEcheMP;                                           // Modification d'échéances
{$IFDEF EAGLCLIENT}
// A FAIRE
{$ELSE}
    38535 : AGLLanceFiche('CP', 'PROFILUSER', '', '', 'FOU');      // Profils utilisateurs
{$ENDIF}
//    38540 : ;// SUPPRIME                                         // Circuit Validation décaissement
    38545 : ExportCFONBBatch(True,False,tslAucun);  // Export CFONB // If LikeS3 viré YMO 26/01/2006
    {$IFDEF CFONB}
    38546 : CP_LanceFicheMulCFONB(''); {Export paramétrable}
    {$ENDIF CFONB}
    38550 : SaisieEffet(OnBqe,srFou);                              // Saisie de trésorerie
    38555 : ReleveFacture;                                         // Relevés de factures
//    7589 : ReleveCompte; // Voir Suivi clients
    38560 : CPLanceFiche_MulEncaDeca ( '', '', 'FLUX=DEC' ) ;                       // Génération des enca/deca
    38565 : CPLanceFiche_MulParamGener ( '', '', 'ACTION=MODIFICATION' ) ;  // Paramétrage des scénarios
    38570 : CP_LanceFicheExport('FOU;');{JP 02/05/05 : FQ 15333}
    38580 : SelectSuiviMP(smpCompenFou);         //FP 21/02/2006
    38585 : GenererCompensation(smpCompenFou);   //FP 21/02/2006
    38590 : CPLanceFiche_GenereEncaDeca( 'FLUX=DEC' ) ;

//======================
//=== Menu Pop Immos ===
//======================
    26520 : CCLanceFiche_MULGeneraux('C;7112000');  // Comptes généraux
    26530 : CPLanceFiche_MULTiers('C;7145000');     // Comptes auxiliaires
    26540 : CPLanceFiche_MULSection('C;7178000');   // Sections analytiques
{$IFDEF AMORTISSEMENT}
    26510 : AMLanceFiche_ListeDesImmobilisations ( '' );
{$ENDIF}

//====================================
//=== MENU NORMALEMENT NON UTILISE ===
//====================================
{$IFNDEF EAGLCLIENT}
//    7278 : Centralisateur(taModif, 'S', FALSE) ; // INUTILISE
//    7302 : Centralisateur(taModif,'U',False) ;   // INUTILISE
//    7318 : Centralisateur(taModif, 'P', FALSE) ; // INUTILISE

// Fin Modif SBO Remaniement des natures d'états...
//    7431 : AGLLanceFiche('CP', 'EPBALBDS', '', 'BDS', 'BDS') ; // Balance générale comparative
//    7432 : AGLLanceFiche('CP', 'EPETASYN', '', 'RC1', 'RLI') ; // Document de synthèse - Charge
//    7433 : AGLLanceFiche('CP', 'EPETASYN', '', 'RP1', 'RLI') ; // Document de synthèse - Produit
//    7452 : AGLLanceFiche('CP', 'EPETASYN', '', 'BPF', 'RLI') ; // Document de synthèse - Passif


    7499 : EncaisseDecaisse(False,'','',True,True,tslAucun) ; // INUTILISE
    7653 : TVAEditeEtat ; // INUTILISE

    (*
    16110 : FicheBanque_AGL('',taModif,0);  // INUTILISE
    16120 : FicheBanqueCP('',taModif,0);    // INUTILISE

    45101 : ParamTable('CPNATURETRESO',taCreat,0,PRien) ;    // INUTILISE
    45102 : ParamTable('CPLUMIEREGENE',taCreat,0,PRien,17) ; // INUTILISE
    45103 : ParamTable('CPLUMIEREAUXI',taCreat,0,PRien,17) ; // INUTILISE
    *)
{$ENDIF} // IFNDEF EAGLCLIENT

//===========================================
//==== Spécificités agricoles 326 ====
//===========================================

{$IFDEF AMORTISSEMENT}
    32610 : AMLanceFiche_SuiviDPI ;
    32620 : AMLanceFiche_Edition ( 'IDD', 'IDD' );
    // BTY 30/10/07 Passage forfait à réel
    32625 : ParamSociete(False,'','SCO_SPECAGRI','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,0);
    32630 : AMLanceFiche_ListeImmosREGFR();
    // BTY 11/07 Ajout édition déclaration
    32635 : AMLanceFiche_Edition ('IFR','IFR' );

{$ENDIF}
     7269  : CPLanceFiche_SaisieTresorerie ;
     else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
   end ;

END ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 10/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Import modif S7 pour S5 Pack avancé
Mots clefs ... : PACK AVANCE
*****************************************************************}
// Modifie les boutons circuits de deca
procedure ChangeMenuDeca ;
var Q      : TQuery ;
    n      : integer ;
    s      : string ;
    StInter: string ;
    sAuto  : string ;
begin
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" ORDER BY CC_CODE',True) ;
n:=0 ;
StInter:='---------------------------------------------------------------------------------------------------' ;
SAuto  :='000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000' ;
While not Q.EOF do
   begin
   Inc(n) ;
   s:=Q.FindField('CC_LIBELLE').AsString ;
   if Trim(s)='' then
      BEGIN
      if Not V_PGI.OutLook then ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+StInter+'" WHERE MN_1=7 AND MN_2=4 AND MN_3=2 AND MN_TAG='+IntToStr(74990+n)) ;
      FMenuG.RemoveItem(74990+n) ;
      END else
      BEGIN
      if Not V_PGI.OutLook then ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+SAuto+'", MN_LIBELLE="'+s+'" WHERE MN_1=7 AND MN_2=4 AND MN_3=2 AND MN_TAG='+IntToStr(74990+n)) ;
      FMenuG.RenameItem(74990+n,s) ;
      END ;
   Q.Next ;
   end ;
Ferme(Q) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/10/2006
Modifié le ... :   /  /    
Description .. : Chargement des modules PCL de l'application
Mots clefs ... : 
*****************************************************************}
procedure SetModulesPCL;
var
  nModules, i : integer;
  a_Modules : array of integer;
  a_Icones : array of integer;
begin
  nModules := 6;
  if VH^.OkModAgricole then Inc (nModules);
  if VH^.OkModImmo then Inc (nModules);
  if VH^.OkModSuiviReg then Inc (nModules,2);
  V_PGI.NbColModuleButtons := nModules div 4;
  if nModules > 7 then
  begin
    if (nModules mod 2) = 0 then V_PGI.NbRowModuleButtons := nModules div 2
    else V_PGI.NbRowModuleButtons := (nModules div 2)+1;
  end
  else V_PGI.NbRowModuleButtons := nModules;
  SetLength(a_Modules, nModules);
  SetLength(a_Icones, nModules);
  a_Modules[0] := 52; a_Icones[0] := 108;   // Traitements courants
  a_Modules[1] := 96; a_Icones[1] := 110;   // Traitements annexes
  i := 2;
  if VH^.OkModImmo then
  begin
    a_Modules[i] := 2;
    a_Icones[i] := 1;
    Inc (i,1);
  end;
  a_Modules[i] := 53; a_Icones[i] := 58;   // Outils Comptables
  Inc (i,1);
  a_Modules[i] := 54; a_Icones[i] := 28;   // Paramètrages
  Inc (i,1);
  a_Modules[i] := 55; a_Icones[i] := 11;   // Missions complémentaires
  Inc (i,1);
  if VH^.OkModSuiviReg then
  begin
    a_Modules[i] := 37; a_Icones[i] := 16;   // suivi fournisseurs
    a_Modules[i+1] := 38; a_Icones[i+1] := 89;   // suivi clients
    Inc (i,2);
  end;
  if (VH^.OkModAgricole) then
  begin
    a_Modules[i] := 326; a_Icones[i] := 87;   // Spécificités agricoles
    Inc (i,1);
  end;
  a_Modules[i] := 56; a_Icones[i] := 49;   // Utilitaires
  FMenuG.SetModules (a_Modules, a_Icones);
end;

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
  Init := True ;
  // Mode pack avancé ?
  VH^.OkModCPPackAvance := True ; // Désormais Par défaut activé // MODIF PACK AVANCE
  // Gestion du multi-lingue ?
  VH^.OkModMultilingue  := False ; // Par défaut non activé // MODIF PACK AVANCE
  // RR dans tous les cas on considère que nous avons cette seria
  VH^.OkModAna := True ;
  VH^.OkModCPPackIFRS := False ;

  if ctxPCL in V_PGI.PGIContexte then
  begin
    VH^.OkModCompta       :=(sAcces[1]='X') ;
    VH^.OkModBudget       :=(sAcces[2]='X') ;
    VH^.OkModImmo         :=(sAcces[3]='X') ;
    VH^.OkModIcc          :=(sAcces[4]='X');
    VH^.OkModRevision     :=(sAcces[5]='X');
    VH^.OkModCre          :=(sAcces[6]='X');
    VH^.OkModGed          :=(sAcces[7]='X');
    VH^.OkModAgricole     :=(sAcces[8]='X');
    VH^.OkModMessagerie   :=(sAcces[9]='X');
    VH^.OkModSuiviReg     :=(sAcces[10]='X');
    VH^.OkModGCD          :=(sAcces[11]='X');
    VH^.OkModDRF          :=(sAcces[12]='X');
    VH^.OkModRIC          :=(sAcces[13]='X');
    VH^.OkModDAS2         :=(sAcces[14]='X');
    VH^.OkModExpertETEBAC :=(sAcces[15]='X');
    VH^.OkModSCAN         :=(sAcces[16]='X');

    { CA - 13/09/2007 : si la compta n'est pas sérialisée, on passe en mode démo
      Demande de MDE }
    if not VH^.OkModCompta then V_PGI.VersionDemo := True;

{$IFDEF GIL}
    VH^.OkModRIC := True;
    VH^.OkModGed := True;
    VH^.OkModICC := True;
{$ENDIF}

  end
  else
  begin
    { En entreprise, ces modules sont toujours accessibles et ne sont pas
      soumis à sérialisation }
    VH^.OkModGCD        := True;
    VH^.OkModDRF        := True;
    VH^.OkModRIC        := True;
    VH^.OkModDAS2       := True;
    {JP 06/07/07 : FQ 18496 : les modules doivent être visibles en version DEMO}
    VH^.OkModCompta:=(sAcces[1]='X') or V_PGI.VersionDemo;
    VH^.OkModBudget:=(sAcces[2]='X') or V_PGI.VersionDemo ;
    {$IFDEF AMORTISSEMENT}
    VH^.OkModImmo:=(sAcces[3]='X')  or V_PGI.VersionDemo;
    VH^.OkModMultilingue  := (sAcces[4]='X') or V_PGI.VersionDemo; // MODIF MULTILINGUE
    VH^.OkModCPPackIFRS   := (sAcces[5]='X') or V_PGI.VersionDemo;  // Modif IFRS
    {$ELSE}
    VH^.OkModImmo:=FALSE ;
    VH^.OkModMultilingue  := (sAcces[3]='X') or V_PGI.VersionDemo; // MODIF MULTILINGUE ?
    VH^.OkModCPPackIFRS   := (sAcces[4]='X') or V_PGI.VersionDemo;  // Modif IFRS
    {$ENDIF}
    If Not VH^.OldTeleTrans Then VH^.OkModEtebac:=TRUE ;
    // Mode S5 pack avance : positionne la série // MODIF PACK AVANCE
    if ( VH^.OkModCPPackAvance ) or ( V_PGI.VersionDemo ) then
    begin
      V_PGI.LaSerie := S7 ;
      // TEST Préférences dans agl pour style du menu :
      V_PGI.ChangeModeOutlook := True ;
    end ;
  end;
  // Si version démo --> Tout accessible mais limité
  if ctxPCL in V_PGI.PGIContexte then SetModulesPCL
  else
  begin
    if V_PGI.VersionDemo then
    begin
      //MENU PCL POUR PGE
      if MenuPCL then
      begin
        if VH^.OkModImmo then
          FMenuG.SetModules([52, 96, 2, 53, 54, 55, 56, 324], [108, 110, 1, 58, 28, 11, 49, 1])
        else
          FMenuG.SetModules([52, 96, 53, 54, 55, 56, 324], [108, 110, 58, 28, 11, 49, 1]);
      end
      else
      begin
        {$IFDEF AMORTISSEMENT}
        V_PGI.NbColModuleButtons := 2;
        V_PGI.NbRowModuleButtons := 6;
        FMenuG.SetModules([9, 11, 13, 16, 361, 14, 12, 20, 8, 17, 18, 324], [108, 78, 81, 90, 104, 58, 22, 1, 11, 28, 49, 1]);
        {$ELSE}
        V_PGI.NbColModuleButtons := 2;
        V_PGI.NbRowModuleButtons := 6;
        FMenuG.SetModules([9, 11, 13, 16, 361, 14, 12, 8, 17, 18, 324], [108, 78, 81, 90, 104, 58, 22, 11, 28, 49, 1]);
        {$ENDIF}
      end;
    end //FIN VERSION DEMO
    else
      if not MenuPCL then
      begin
        if not VH^.OkModCompta then
        begin
          begin
            FMenuG.SetModules([17, 18], [28, 49]);
            V_PGI.NbColModuleButtons := 1;
          end;
        end
        else
        begin // Nécessairement compta + éventuels autres modules
          if VH^.OkModBudget then
          begin
            V_PGI.NbColModuleButtons := 2;
            V_PGI.NbRowModuleButtons := 6;
            if VH^.OkModImmo then
              FMenuG.SetModules([9, 11, 13, 16, 361, 14, 12, 20, 8, 17, 18, 324], [108, 78, 81, 90, 104, 58, 22, 1, 11, 28, 49, 1])
            else
              FMenuG.SetModules([9, 11, 13, 16, 361, 14, 12, 8, 17, 18, 324], [108, 78, 81, 90, 104, 58, 22, 11, 28, 49, 1]);
          end
          else
          begin
            if VH^.OkModImmo then
            begin
              V_PGI.NbColModuleButtons := 2;
              V_PGI.NbRowModuleButtons := 6;
              FMenuG.SetModules([9, 11, 13, 16, 361, 14, 20, 8, 17, 18, 324], [108, 78, 81, 90, 104, 58, 1, 11, 28, 49, 1]);
            end
            else
            begin
              V_PGI.NbColModuleButtons := 2;
              V_PGI.NbRowModuleButtons := 5;
              FMenuG.SetModules([9, 11, 13, 16, 361, 14, 8, 17, 18, 324], [108, 78, 81, 90, 104, 58, 11, 28, 49, 1]);
            end;
          end;
        end;
      end
      else //MODE MENUPCL POUR PGE
      begin
        if MenuPCL then
        begin
          if VH^.OkModImmo then
            FMenuG.SetModules([52, 96, 2, 53, 54, 55, 56], [108, 110, 1, 58, 28, 11, 49])
          else
            FMenuG.SetModules([52, 96, 53, 54, 55, 56], [108, 110, 58, 28, 11, 49]);
        end;
      end;
   end;
{$IFDEF TT}
  VH^.OkModGed:= true ; VH^.OkModGCD := true ;
{$ENDIF}
{$IFDEF SCANGED}
  if not VH^.OkModGed then OnGlobalAfterImportGED := nil;
{$ENDIF}
END ;

procedure AssignLesZoom ( LesImmos : boolean ) ;
BEGIN
  {$IFDEF AMORTISSEMENT}
if LesImmos then
   BEGIN
   ProcZoomEdt:=imZoomEdtEtatImmo ;
   ProcCalcEdt:=CalcOLEEtatImmo ;
   END else
   {$ENDIF}
   BEGIN
   ProcZoomEdt:=ZoomEdtEtat ;
   ProcCalcEdt:=CalcOLEEtat ;
   END ;

//{$ENDIF}
END ;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
  UpdateSeries ;
  AssignLesZoom((NumModule=20) or (NumModule=2)) ;

  Case NumModule of
    8,9,11,13,14,16,17,18,52,53,54,55,56,96,324 : ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ; // FQ 10360
    12 : ChargeMenuPop(integer(hm3),FMenuG.DispatchX) ;
// BTY 07/05 Fiche 14754 Supprimer les anciennes éditions en cwas
    20,2 :
    begin
       // MVG 05/07/2007 FQ  20896 hm3 => hm2
       ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ;
       // fq 18008 - mbo 31.05.2006 cacher anciennes éditions dans tous les cas{$IFDEF EAGLCLIENT}
       if  (NumModule = 20) or (NumModule = 2) then
            FMenuG.RemoveGroup(20400,True) ;
       // {$ENDIF}
    end;
// fin
    else ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
  END ;

  // Tripotage des menus
  TripotageMenu (NumModule);

  // Change les captions des menus Tiers / Décaissements
  if NumModule=11 then ChangeMenuDeca ;

  if FMenuG.PopupMenu=Nil then FMenuG.PopUpMenu:=ADDMenuPop(Nil,'','') ;

  FMDispS5.NumModule:=NumModule ;

  if (NumModule = 17)   then         // GP N° 5594
    FMenuG.RemoveItem(1272) ;

  // ajout me pour ca3
  if (NumModule = 96) then FMenuG.RemoveGroup(7637,True) ;
  if (NumModule = 14)   then FMenuG.RemoveGroup(-7660,True) ;
END ;

function ChargeFavoris : boolean ;
begin
  if ctxPCL in V_PGI.PGIContexte
    then AddGroupFavoris( FMenuG , '54;2;55;96')
    else AddGroupFavoris( FMenuG , '12;20;8;17;18;324') ;
  result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/09/2003
Modifié le ... :   /  /
Description .. : Suppression d'une ligne ou d'un groupe de menu avec mise
Suite ........ : à jour de la liste des tag supprimés
Mots clefs ... :
*****************************************************************}
procedure RemoveFromMenu ( const iTag : integer; const bGroup : boolean; var stAExclure : string );
begin
  if bGroup then FMenuG.RemoveGroup (iTag, True)
  else FMenuG.RemoveItem (iTag);
  stAExclure := stAExclure+IntToStr(iTag)+';';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/09/2003
Modifié le ... : 06/09/2006
Description .. : - Suppression des menus
Suite ........ : - Constitution de la liste des tag à supprimer pour la gestion
Suite ........ : des favoris
Suite ........ : - LG - 06/09/2006 - suppression du menu tva si le 
Suite ........ : paramsoc n'est pas present
Mots clefs ... : 
*****************************************************************}
procedure TripotageMenuModePCL ( const TNumModule : array of integer; var stAExclure : string );
var i, NumModule : integer;
begin
  for i := 0 to High(TNumModule) do
  begin
    NumModule := TNumModule[i];
    case NumModule of
      52 : begin
             if not GetParamSocSecur('SO_CPPCLSAISIEQTE', False) then
               RemoveFromMenu(7417, False, stAExclure);
             if not GetParamSocSecur('SO_CPPCLSAISIETVA', False) then
               RemoveFromMenu(52270, False, stAExclure);

             if (not VH^.OkModRIC) or (VH^.Revision.Plan = '') then
               RemoveFromMenu(52150, True, StAExclure)
             else
             begin
               if JaiLeRoleCompta( rcSuperviseur ) then  // N3 ou N4
                 RemoveFromMenu(52152, False, StAExclure)
               else
                 if JaiLeRoleCompta( rcReviseur ) then // N2
                   RemoveFromMenu(52153, False, StAExclure)
                 else
                   if JaiLeRoleCompta( RcCollaborateur ) then // N1
                   begin
                     RemoveFromMenu(52152, False, StAExclure);
                     RemoveFromMenu(52153, False, StAExclure);
                   end;

             end;
           end;
      { Outils comptables }
      53 :  begin
              { Suppression des éditions légales }
              RemoveFromMenu (7709, False, stAExclure);
              //06/12/2006 YMO Norme NF 203

              if (GetParamSocSecur('SO_CPANNULCLO', False)>0)
{$IFNDEF CERTIFNF}
              or (GetParamSocSecur('SO_CPCONFORMEBOI', False))
{$ENDIF}
              then
              begin
                RemoveFromMenu(7694, False, stAExclure);
                RemoveFromMenu(7703, False, stAExclure);
                RemoveFromMenu(7707, False, stAExclure);
              end;
                                                                    
{$IFNDEF CERTIFNF}
              if GetParamSocSecur('SO_CPCONFORMEBOI', False) then
{$ENDIF}              
              RemoveFromMenu(7737, False, stAExclure);
            end;
      { Paramétrages }
      54 :  begin
              if GetParamSocSecur('SO_CPPCLSANSANA','-',True) then RemoveFromMenu (54600, True, stAExclure);
              { Réception ETEBAC }
              if (not VH^.OldTeleTrans) then RemoveFromMenu (1732, False, stAExclure);

              // Suppression des menus d'accès aux plans de rupture
              RemoveFromMenu(1375, False, StAExclure);
              RemoveFromMenu(1380, False, StAExclure);
              RemoveFromMenu(1395, False, StAExclure);
              RemoveFromMenu(1400, False, StAExclure);
              RemoveFromMenu(-10, False, StAExclure);
              RemoveFromMenu(-14, False, StAExclure);

              // Suppression du nouvel écran de la révision si pas sérialisé
              if (not VH^.OkModRic) then
                RemoveFromMenu(54725, False, StAExclure);

            end;
      { Missions complémentaires }
      55 :  begin
              { Si Budget pas sérialisé, le menu Budget n'apparaît pas ! }
              if VH^.OkModBudget=False then RemoveFromMenu (55300, True, stAExclure);
              {JP 23/04/07 : FQ 15945 : des menus analytiques}
              if GetParamSoc('SO_CPPCLSANSANA') then begin
                RemoveFromMenu (7361, True, stAExclure);
                RemoveFromMenu (15213, False, stAExclure);
                RemoveFromMenu (15217, False, stAExclure);
                RemoveFromMenu (15219, False, stAExclure);
                RemoveFromMenu (15225, False, stAExclure);
                RemoveFromMenu (15227, False, stAExclure);
                RemoveFromMenu (15135, False, stAExclure);
                RemoveFromMenu (15139, False, stAExclure);
                RemoveFromMenu (15145, False, stAExclure);
                RemoveFromMenu (15137, False, stAExclure);
                RemoveFromMenu (15322, False, stAExclure);
                RemoveFromMenu (15323, False, stAExclure);
                RemoveFromMenu (15324, False, stAExclure);
                RemoveFromMenu (15332, False, stAExclure);
                RemoveFromMenu (15333, False, stAExclure);
                RemoveFromMenu (15334, False, stAExclure);
                RemoveFromMenu (15284, False, stAExclure);
                RemoveFromMenu (1528 , False, stAExclure);
                RemoveFromMenu (1522 , False, stAExclure);
                RemoveFromMenu (15146, False, stAExclure);
              end;
              {$IFDEF EAGLCLIENT}
              // RemoveFromMenu (7381, False, stAExclure) ; // Ré-imputations
              RemoveFromMenu (-72, False, stAExclure) ; // Encaissements
              RemoveFromMenu (-73, False, stAExclure) ; // Décaissements
              RemoveFromMenu (7529, False, stAExclure) ; // Modification d'échéances
              //YMO 03/07/2006 FQ 17850 Menu Estimation des coûts financiers dispo en CWAS
//              RemoveFromMenu (7598, False, stAExclure) ; // Estimation des coûts financiers
              RemoveFromMenu (52701, False, stAExclure) ; // Balance âgée (***)
              RemoveFromMenu (7599, False, stAExclure) ; // Bordereaux d'accompagnement
              RemoveFromMenu (7596, False, stAExclure) ; // Réédition de traites clients
              RemoveFromMenu (7597, False, stAExclure) ; // Réédition de BOR fournisseurs
              RemoveFromMenu (-102, False, stAExclure) ; // Export CFONB
              RemoveFromMenu (1485, False, stAExclure) ; // Etapes d'encaiss
              RemoveFromMenu (1490, False, stAExclure) ; // Etapes de décaissements
              RemoveFromMenu (1491, False, stAExclure) ; // Circuits de validation des décaissements
              RemoveFromMenu (1235, False, stAExclure) ; // Paramétrage des lettres de relance
              RemoveFromMenu (1240, False, stAExclure) ; // Paramétrage des relevés de facture
              RemoveFromMenu (1245, False, stAExclure) ; // Paramétrage des relevés de compte
              //RemoveFromMenu (1247, False, stAExclure) ; // Paramétrage des éditions en cours de saisie
              //RemoveFromMenu (1248, False, stAExclure) ; // Paramétrage des bons à payer
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then RemoveFromMenu (1267, False, stAExclure) ; // Paramétrage des lettre chèques
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then RemoveFromMenu (1268, False, stAExclure) ; // Paramétrage des traites LCR/BOR
              //RemoveFromMenu (1257, False, stAExclure) ; // Paramétrage des bordereaux d'accompagnement
              {$ENDIF}
              { Bilan analytique }
              RemoveFromMenu (7449, False, stAExclure);
              { Télétransmission ETEBAC }
              if (not VH^.OldTeleTrans) then RemoveFromMenu (7594, False, stAExclure);

            end;
      { Utilitaires }
      56 :  begin
              { Modification d'entête de pièces }
              RemoveFromMenu (56180, False, stAExclure);
              {$IFDEF EAGLCLIENT}
              { Suppression du menu "contrôles" en Web Access }
                  {$IFDEF CONTREPARTIE}
                      // ajout me
                       FMenuG.RemoveItem(3210);
                       FMenuG.RemoveItem(3220);
                       FMenuG.RemoveItem(3185);
                  // fin ajout me
                  {$ELSE}
                        RemoveFromMenu (56200, True , stAExclure);
                  {$ENDIF CONTREPARTIE}
              { Suppression du menu "purge" en Web Access }
              RemoveFromMenu (56190, False , stAExclure);
              {$ENDIF}
              Init := false ;
              //06/12/2006 YMO Norme NF 203
{$IFNDEF CERTIFNF}
              // On supprime l'acces au mul du journal des evenements
              RemoveFromMenu (56205, False, stAExclure) ;
              if GetParamSocSecur('SO_CPCONFORMEBOI', False) then
{$ENDIF}
               RemoveFromMenu(3245, False, stAExclure);
            end;
      { Spécificités agricoles }
      326 : begin
                 RemoveFromMenu (-32631, TRUE , stAExclure); // FQ 19234
                 if not  CPEstComptaAgricole then
                 begin
                      RemoveFromMenu (-32611, TRUE , stAExclure); // FQ 19234
                      RemoveFromMenu (-32621, TRUE , stAExclure);// FQ 19234
                 end;
            end;
      { Traitements annexes }
      96 :  begin
              { Si ICC pas sérialisé, le menu ICC n'apparaît pas }
              if VH^.OkModIcc=False then RemoveFromMenu (96100, True, stAExclure);
              { Si CRE pas sérialisé, le menu CRE n'apparaît pas }
              if VH^.OkModCre=False then RemoveFromMenu (96200, True, stAExclure);
              { Si Immo pas sérialisé, le menu Immo n'apparaît pas }
              if VH^.OkModImmo=False then RemoveFromMenu (20, True, stAExclure);
              { Si GCD pas sérialisé, le menu GCD n'apparaît pas }
              if VH^.OkModGCD=False then RemoveFromMenu (96400, True, stAExclure);
              { Si DAS2 pas sérialisé, le menu DAS2 n'apparaît pas }
              if VH^.OkModDAS2=False then RemoveFromMenu (96500, True, stAExclure);
              { Réception ETEBAC }
              if (not VH^.OldTeleTrans) then RemoveFromMenu ( 7772, False, stAExclure);
              if (GetParamsoc ('SO_CPLIENGAMME') <> 'S1')
              // and (GetParamSocSecur ('SO_CPSYNCHROSX', '') <> TRUE)  CA - 18/09/2006 : peu importe le mode.
              then
              begin
                        FMenuG.RemoveItem(96323);
                        FMenuG.RemoveItem(96324);
              end;
              {JP 05/03/07 : Mise en place du nouveau pointage}
              if EstSpecif('51213') and (ctxPcl in V_PGI.PGIContexte) then begin
                RemoveFromMenu(7603, False, stAExclure); {Relevés}
                RemoveFromMenu(7605, False, stAExclure); {Mouvements bancaires}
              end else begin
                 {En Mode PCL}
                 RemoveFromMenu(-96120, True, stAExclure); {Rappro}
                 RemoveFromMenu(-96120, False, stAExclure); {Rappro}
                 {Menu d'intégration : Demande du 19/04/07}
                 RemoveFromMenu(-106, True, stAExclure); {Rappro}
                 RemoveFromMenu(-106, False, stAExclure); {Rappro}
              end;
              { Appel de la fiche établissement de la paie si gestion des établissements }
              if not GetParamSocSecur('SO_ETABLISCPTA',False) then
                RemoveFromMenu(96590, False, stAExclure); {Rappro}
            end;
    end;
  end;
end;

function IsOracle : Boolean;
begin
  Result := (V_PGI.Driver=dbOracle7) or (V_PGI.Driver=dbOracle8);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Gestion des modules Internationnaux
Mots clefs ... :
*****************************************************************}
procedure TripotageMenuInternat ( const TNumModule : array of integer; var stAExclure : string );
var i, NumModule : integer;
begin

  for i := 0 to High(TNumModule) do
  begin
    NumModule := TNumModule[i];
    case NumModule of
      // ========================
      // === Reporting et TVA ===
      // ========================
      14 : begin
              // Belgique
              if QuelPaysLocalisation <> CodeIsoDuPays('BEL') then
              begin
                 // Module de déclaration TVA
                 RemoveFromMenu (-1492, False, stAExclure) ; //SDA le 12/12/2007 devient un sous-menu
              end;
           end;
      17 : begin
            {$IFNDEF ADRESSE}
              //Gestion des adresses
              RemoveFromMenu (1396, False, stAExclure) ;
              RemoveFromMenu (1397, False, stAExclure) ;
           {$ENDIF}
           {$IFNDEF TAXES}
             RemoveFromMenu (-50, False, stAExclure) ;
             RemoveFromMenu (1770, False, stAExclure) ;
             RemoveFromMenu (1775, False, stAExclure) ;
             RemoveFromMenu (1780, False, stAExclure) ;
             RemoveFromMenu (1790, False, stAExclure) ;
             RemoveFromMenu (1795, False, stAExclure) ;
             RemoveFromMenu (1800, False, stAExclure) ;
            {$ENDIF}
           end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 12/09/2003
Modifié le ... : 22/09/2003
Description .. : - Suppression des menus PGE
Suite ........ : - Constitution de la liste des tag à supprimer pour la gestion
Suite ........ : des favoris PGE
Mots clefs ... :
*****************************************************************}
procedure TripotageMenuModePGE ( const TNumModule : array of integer; var stAExclure : string );
var i, NumModule : integer;
begin
  // Autres modules à exclures
  for i := 0 to High(TNumModule) do
  begin
    NumModule := TNumModule[i];
    case NumModule of

      // ===============
      // === ANALYSE ===
      // ===============
      8 :  begin
           if not EstComptaIFRS then
             begin
             RemoveFromMenu (8315, False, stAExclure) ;
             RemoveFromMenu (8515, False, stAExclure) ;
             end ;
           //SG6 17/12/2004 Si Mode analytique croisaxe > on cache le menu analyse multi-axes
           if VH^.AnaCroisaxe then
              RemoveFromMenu(-8205,True,stAExclure);
           end;

      // =================
      // === Ecritures ===
      // =================
      9 :  begin
           {$IFDEF EAGLCLIENT}
//Lek190106             RemoveFromMenu (7268, False, stAExclure) ; // Brouillard
//Lek190106             RemoveFromMenu (7289, False, stAExclure) ; // Brouillard
//CA260407             RemoveFromMenu (7381, False, stAExclure) ; // Ré-imputations
//Lek190106             RemoveFromMenu (7385, False, stAExclure) ; // Brouillard
//Lek190106             RemoveFromMenu(7670, False, stAExclure) ; // Révision : Brouillard
             RemoveFromMenu(7682, False, stAExclure) ; // Révision :  Mode révision
             RemoveFromMenu(52650, True, stAExclure) ; // Autres éditions Lek 190106
           {$ENDIF}
           // Ecriture IFRS
           if not EstComptaIFRS then
             RemoveFromMenu (-9200, True, stAExclure) ; // Saisie IFRS
//thl           RemoveFromMenu (7347, True, stAExclure) ; // Engagement Lek 270406
           end ;

      // =============
      // === Tiers ===
      // =============
      11 : begin
           {$IFDEF EAGLCLIENT}
             RemoveFromMenu (7496,  True, stAExclure) ;   // Encaissements
             RemoveFromMenu (-74999, True, stAExclure) ;   // Décaissements

             // RemoveFromMenu (7520,   False, stAExclure) ;  // Lettrage
             // RemoveFromMenu (7523,   False, stAExclure) ;  // Lettrage
             RemoveFromMenu (7529,   False, stAExclure) ;  // Lettrage modif d'échéance
             //YMO 03/07/2006 FQ 17850 Menu Estimation des coûts financiers dispo en CWAS
//             RemoveFromMenu (7598,   False, stAExclure) ;   // Estimation du coût des tiers

             // Editions de suivi :
// Lek 190106             RemoveFromMenu (7535,   False, stAExclure) ;   // Justif de solde
// Lek 190106            RemoveFromMenu (7550,   False, stAExclure) ;   // GL Agé
// Lek 190106            RemoveFromMenu (7559,   False, stAExclure) ;   // GL Ventilé
           //b Lek 11/10/05 activer Tiers payeurs
//             RemoveFromMenu (-7230,  True, stAExclure) ;   // Tiers Payeurs
//             RemoveFromMenu (7234,   False, stAExclure) ;   //  : BalanceAge2TP ;
//             RemoveFromMenu (7236,   False, stAExclure) ;   //  : GLivreAgeTP ;
//             RemoveFromMenu (7235,   False, stAExclure) ;   //  : BalVentileTP ;
//             RemoveFromMenu (7237,   False, stAExclure) ;   //  : GLVentileTP ;
           //e Lek 11/10/05



             RemoveFromMenu (-52700, True, stAExclure) ;   // Autres éditions

           {$ELSE} {Lek 111005 Prêt à enlever quand l'édition est validée}
//             RemoveFromMenu (52710,   False, stAExclure) ;  // : BrouillardTP ;{Lek&FP 071005}
//             RemoveFromMenu (52715,   False, stAExclure) ;  // : BalanceAge2TP ;{Lek&Fréd 111005}
//             RemoveFromMenu (52720,   False, stAExclure) ;  // : GLivreAgeTP ;
//             RemoveFromMenu (52725,   False, stAExclure) ;  // : BalVentileTP ;
//             RemoveFromMenu (52730,   False, stAExclure) ;  // : GLVentileTP ;

           {$ENDIF}
           // Cacher le module d'écart de conversion du lettrage
           RemoveFromMenu (7524, False, stAExclure) ; // Lettrage --> Ecart de conversion
           end ;

      // ===============
      // === BUDGETS ===
      // ===============
      12 : begin
             {JP 23/04/07 : Vieux états, on les cache aussi en 2/3}
             RemoveFromMenu (15350, False, stAExclure); // Etats libres
             RemoveFromMenu (52800, True, stAExclure); // Autres éditions Lek190106
           end ;

      // ================
      // === EDITIONS ===
      // ================
      13 : begin
           {$IFDEF EAGLCLIENT}
             // Journaux
//Lek190106             RemoveFromMenu (7409, False, stAExclure);   // Périodique par général
             // Editions sur tables libres ???
             if (not (EstSerie(S7))) then
               begin
               RemoveFromMenu (7419, False, stAExclure);   // Table libre uniquement pour S7
               RemoveFromMenu (7429, False, stAExclure);   // Table libre uniquement pour S7
               end ;
             // Etats de synthèses
             RemoveFromMenu (7444, False, stAExclure);   // Etats chainés
             // Autres éditions
             RemoveFromMenu (-52580, True, stAExclure);  // Autres éditions
             // Autres
             RemoveFromMenu (7490, False, stAExclure); // Récapitulatifs des éditions
           {$ENDIF}
           // BPY le 21/06/2004 => Restitutions multi sociétés
           RemoveFromMenu(8623,false,stAExclure);    // Grand Livre Analytique
           RemoveFromMenu(8627,false,stAExclure);    // Grand livre Général par analytique
           RemoveFromMenu(8628,false,stAExclure);    // Grand livre Analytique par général
           // Fin BPY
           end;

      // ========================
      // === Reporting et TVA ===
      // ========================
      14 : begin
           if VH^.PaysLocalisation=CodeISOES then
             begin
             //Gestion TVA
             RemoveFromMenu(7646, FALSE, stAExclure) ; // Modification des Bases HT
             RemoveFromMenu(7648, FALSE, stAExclure) ; // Régularisation des acomptes
             //Editions TVA
             RemoveFromMenu(7650, FALSE, stAExclure) ; // Gestion directe des factures
             RemoveFromMenu(7651, FALSE, stAExclure) ; // Gestion par etapes des factures
             RemoveFromMenu(7652, FALSE, stAExclure) ; // Etat des acomptes
             RemoveFromMenu(7654, FALSE, stAExclure) ; // Validation des editions de TVA
             end
             //SDA le 12/12/2007 ces éléments ne concernent que l'Espagne - la France et les autres
             //pays NON
             //else if VH^.PaysLocalisation=CodeISOFR then
             else if VH^.PaysLocalisation<>CodeISOES then //SDA le 12/12/2007
             begin
             RemoveFromMenu(140501, FALSE, stAExclure) ; // Brouillon Modèle 347
             RemoveFromMenu(-140502, FALSE, stAExclure) ; // Livre TVA
             RemoveFromMenu(-140500,TRUE, stAExclure) ; // Modèles officiaux ESPAGNOLS impôts
             end ; //XVI 24/02/2005
           end ;

      // ==================
      // === Traitement ===
      // ==================
      16 : begin
           {$IFDEF EAGLCLIENT}
             RemoveFromMenu (77120, False, stAExclure);  // Clotures / validations -> Etats des éditions légales
             RemoveFromMenu (-7625, True, stAExclure);   // Analyse de trésorie
             If (Not VH^.OldTeleTrans) Then
               begin                                                                    // NE PAS SUPPRIMER CETTE LIGNE
               RemoveFromMenu (7772, False, stAExclure) ; // Réception Etebac 3         // NE PAS SUPPRIMER CETTE LIGNE
               RemoveFromMenu (7594, False, stAExclure) ; // Télétransmission Etebac    // NE PAS SUPPRIMER CETTE LIGNE
               end;                                                                     // NE PAS SUPPRIMER CETTE LIGNE
             // Sur comptes
             RemoveFromMenu (7719, False, stAExclure) ; // Affectations sur TL écritures
             // Pointage
             // Validations
//Lek 190106     RemoveFromMenu (7709, False, stAExclure) ; // Editions légales
             // Clôtures
//Lek 190106             RemoveFromMenu (7757, False, stAExclure) ;
//Lek 190106             RemoveFromMenu (7760, False, stAExclure) ;
//Lek 190106             RemoveFromMenu (7766, False, stAExclure) ;
//Lek 190106             RemoveFromMenu (7769, False, stAExclure) ;
             RemoveFromMenu (52860, True, stAExclure) ; {Lek190106 Autres éditions}
             RemoveFromMenu (7683, False, stAExclure) ;
             // Relevés
             RemoveFromMenu (7772, False, stAExclure) ;
             RemoveFromMenu (1755, False, stAExclure); // Partage du référentiel
           {$ENDIF}
             {JP 17/08/05 : FQ 15855 : suppression du menu dans la SocRef 704 du 25/08/05
             RemoveFromMenu(7607, False, stAExclure ) ; // Dépointage}

             //06/12/2006 YMO Norme NF 203
             if (GetParamSocSecur('SO_CPANNULCLO', False)>0)
{$IFNDEF CERTIFNF}
                or (GetParamSocSecur('SO_CPCONFORMEBOI', False))
{$ENDIF}
             then
             begin
                RemoveFromMenu(7694, False, stAExclure);
                RemoveFromMenu(7703, False, stAExclure);
                RemoveFromMenu(7707, False, stAExclure);
             end;

{$IFNDEF CERTIFNF}
             if GetParamSocSecur('SO_CPCONFORMEBOI', False) then
{$ENDIF}
                RemoveFromMenu(7737, False, stAExclure);

               {JP 02/10/07 : FQ TRESO 10517 : 7619, c'est l'état de rappro et non l'intégration
                RemoveFromMenu(7619, False, stAExclure);
                JP 16/10/07 : FQ 21656 : on laisse le menu disponible pour le moment pour la saisie
                              de trésorerie
             if EstPointageCache then
               RemoveFromMenu(7601, False, stAExclure); {Intégration des Relevés}

             {JP 05/03/07 : Mise en place du nouveau pointage}
             if EstSpecif('51213') and (ctxPcl in V_PGI.PGIContexte) then begin
               RemoveFromMenu(7603, False, stAExclure); {Relevés}
               RemoveFromMenu(7605, False, stAExclure); {Mouvements bancaires}
             end else begin
               {En mode PGE}
               If (ctxPcl in V_PGI.PGIContexte) Then Else
                 BEGIN
                 If OkNewPointage Then
                   BEGIN
                   RemoveFromMenu(-7772, True , stAExclure); {Menu rapprochement}
                   RemoveFromMenu(16110, True , stAExclure); {Activation "pointage avancé"}
                   END Else
                   BEGIN
                   RemoveFromMenu(7603, True , stAExclure); {Menu &Relevés}
                   RemoveFromMenu(7601, True , stAExclure); // &Intégration des relevés
                   RemoveFromMenu(7605, True , stAExclure); // &Opérations bancaires
                   END ;
                 END ;
             end;
           end;

      // ================================
      // === Structures et paramètres ===
      // ================================
      17 : begin
           {$IFDEF EAGLCLIENT}
             // Paramètres //RR 02/06/2003
             RemoveFromMenu(1315, False, stAExclure ) ; // Plans de référence

             {JP 30/09/05 : Je pense qu'il s'agissait d'un oubli
             RemoveFromMenu(-9, False, stAExclure ) ;   // Comptes de correspondance}

             RemoveFromMenu(-10, False, stAExclure ) ;  // Plans de ruptures des éditions
             RemoveFromMenu(-12, False, stAExclure ) ;  // Etapes de réglement
             // Banques
             RemoveFromMenu(1730, False, stAExclure ) ; // Codes résidents étrangers
             RemoveFromMenu(1272, False, stAExclure ) ; // Paramètrage
             // Documents
             //RemoveFromMenu (1225, True, stAExclure) ;
             // RemoveFromMenu (-1280, False, stAExclure) ; // Edition paramétrables NE PAS SUPPRIMER CETTE LIGNE
              RemoveFromMenu (1235, False, stAExclure) ; // Paramétrage des lettres de relance
              RemoveFromMenu (1240, False, stAExclure) ; // Paramétrage des relevés de facture
              RemoveFromMenu (1245, False, stAExclure) ; // Paramétrage des relevés de compte
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then RemoveFromMenu (1267, False, stAExclure) ; // Paramétrage des lettre chèques
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then RemoveFromMenu (1268, False, stAExclure) ; // Paramétrage des traites LCR/BOR

              {JP 14/01/08 : Reprise de ce qui était fait dans MDispMp.Pas}
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then
                RemoveFromMenu (1266, False, stAExclure); // Lettres prélèvements
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then
                RemoveFromMenu (1267, False, stAExclure);
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then
                RemoveFromMenu (1268, False, stAExclure);
              if not GetParamSocSecur('SO_CPDOCAVECETAT', True, True) then
                RemoveFromMenu (1269, False, stAExclure); // Lettres virements
           {$ENDIF}

           {JP 30/09/05 : FQ 16569 : Plan de correspondance Généraux}
           if not GetParamSocSecur('SO_CORSGE1', True, True) and
              not GetParamSocSecur('SO_CORSGE2', True, True) then RemoveFromMenu(1325, False, stAExclure);
           {JP 30/09/05 : FQ 16569 : Plan de correspondance Auxiliaire}
           if not GetParamSocSecur('SO_CORSAU1', True, True) and
              not GetParamSocSecur('SO_CORSAU2', True, True) then RemoveFromMenu(1330, False, stAExclure);
           {JP 30/09/05 : FQ 16569 : Plan de correspondance Axe analytique 1}
           if not GetParamSocSecur('SO_CORSA11', True, True) and
              not GetParamSocSecur('SO_CORSA12', True, True) then RemoveFromMenu(1345, False, stAExclure);
           {JP 30/09/05 : FQ 16569 : Plan de correspondance Axe analytique 2}
           if not GetParamSocSecur('SO_CORSA21', True, True) and
              not GetParamSocSecur('SO_CORSA22', True, True) then RemoveFromMenu(1350, False, stAExclure);
           {JP 30/09/05 : FQ 16569 : Plan de correspondance Axe analytique 3}
           if not GetParamSocSecur('SO_CORSA31', True, True) and
              not GetParamSocSecur('SO_CORSA32', True, True) then RemoveFromMenu(1355, False, stAExclure);
           {JP 30/09/05 : FQ 16569 : Plan de correspondance Axe analytique 4}
           if not GetParamSocSecur('SO_CORSA41', True, True) and
              not GetParamSocSecur('SO_CORSA42', True, True) then RemoveFromMenu(1360, False, stAExclure);
           {JP 30/09/05 : FQ 16569 : Plan de correspondance Axe analytique 5}
           if not GetParamSocSecur('SO_CORSA51', True, True) and
              not GetParamSocSecur('SO_CORSA52', True, True) then RemoveFromMenu(1365, False, stAExclure);

           {JP 30/09/05 : FQ 16569 : Plan de correspondance des Budgets.
            Semble inutile car jamais utilisé : vu avec OG et EC
           if not GetParamSocSecur('SO_CORSBU1', True, True) and
              not GetParamSocSecur('SO_CORSBU2', True, True) then}
           RemoveFromMenu(1426, False, stAExclure);

           {JP 18/10/05 : FQ 16900 : Si aucun plan de correspondance, on cache le groupe de menu}
           if not GetParamSocSecur('SO_CORSGE1', True, True) and not GetParamSocSecur('SO_CORSGE2', True, True) and
              not GetParamSocSecur('SO_CORSAU1', True, True) and not GetParamSocSecur('SO_CORSAU2', True, True) and
              not GetParamSocSecur('SO_CORSA11', True, True) and not GetParamSocSecur('SO_CORSA12', True, True) and
              not GetParamSocSecur('SO_CORSA21', True, True) and not GetParamSocSecur('SO_CORSA22', True, True) and
              not GetParamSocSecur('SO_CORSA31', True, True) and not GetParamSocSecur('SO_CORSA32', True, True) and
              not GetParamSocSecur('SO_CORSA41', True, True) and not GetParamSocSecur('SO_CORSA42', True, True) and
              not GetParamSocSecur('SO_CORSA51', True, True) and not GetParamSocSecur('SO_CORSA52', True, True) then
             RemoveFromMenu(-9, False, stAExclure);

           RemoveFromMenu (-1426, False, stAExclure); // Compte de regroupement
           if not EstComptaTreso then
             begin
             RemoveFromMenu (1706, False, stAExclure); // Agences bancaires
             RemoveFromMenu (1707, False, stAExclure); // Comptes bancaires
             end;
            // YMO 28/02/2006 FQ16919 la suppression du menu 1755 (partage du référentiel) était systématique
           if not EstMultiSoc then
             begin
             RemoveFromMenu (1755, False, stAExclure) ;
             RemoveFromMenu (1765, False, stAExclure) ;
             end ;
           RemoveFromMenu(54732, False, stAExclure);
           end ;


      // ================================
      // === Outils et administration ===
      // ================================
      18 : begin
           //Anomalie d'écriture
           if not V_PGI.SAV then
             RemoveFromMenu(3212, False, stAExclure);
           if not EstComptaTreso then
             RemoveFromMenu (3217, False, stAExclure); {JP 10/03/04 : mise ne place de l'initialisation des agences}
           if not VH^.OkModMultilingue
             then RemoveFromMenu (3100, True, stAExclure); // Menu installation // 14693

           {$IFDEF EAGLCLIENT}
             // Société
             FMenuG.RemoveGroup(3125,True) ;
             RemoveFromMenu (-1426, False, stAExclure); // Compte de regroupement
             // Utilisateurs et accès
         //  RemoveFromMenu (3185, False, stAExclure); // Suivi activité FQ17895 YMO 27/06/06 Remise en place en CWAS
             // Outils
             RemoveFromMenu (3210, False, stAExclure);
             {$IFDEF CONTREPARTIE}
             {$ELSE}
             RemoveFromMenu (3215, False, stAExclure);
             {$ENDIF}
             RemoveFromMenu (3220, False, stAExclure);
             RemoveFromMenu (3221, False, stAExclure);

             // YM 13/09/2005 FQ 16586 Menu remis en mode Web acces
             // module Administration & outils - menu Outils - commande Changement du mode analytique
             // RemoveFromMenu (3227, False, stAExclure); //Changement du mode analytique

           {$ENDIF}
           //06/12/2006 YMO Norme NF 203           
{$IFNDEF CERTIFNF}
           // Suppression de l'acces au journal des evenements
           RemoveFromMenu (3169, False, stAExclure) ;
           if GetParamSocSecur('SO_CPCONFORMEBOI', False) then
{$ENDIF}           
               RemoveFromMenu(3245, False, stAExclure);
           // GP le 21/08/2008 : A virer. fiche N° 22060
           RemoveFromMenu (3227, False, stAExclure); //Changement du mode analytique

           end;

      // =====================
      // === Amortissement ===
      // =====================
      20 : Begin
           {$IFDEF EAGLCLIENT}
             FMenuG.RemoveGroup(20400,True);
           {$ENDIF}
           end ;
      361 : begin
              if not GetParamSocSecur('SO_ETABLISCPTA',False) then
                RemoveFromMenu(96590, False, stAExclure); {Rappro}
              // CA - 16/10/2007 - FQ 21647 - Pas utilisé pour l'instant, on supprime
              RemoveFromMenu (-96325,False, stAExclure);
              RemoveFromMenu (96325,False, stAExclure);
              RemoveFromMenu (96326,False, stAExclure); 
            end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/09/2003
Modifié le ... : 03/09/2004
Description .. : - Suppression des menus du module amortissement
Suite ........ : - Constitution de la liste des tag à supprimer pour la gestion
Suite ........ : des favoris
Mots clefs ... :
*****************************************************************}
procedure TripotageMenuAmortissement ( var stAExclure : string );
begin
  { Suppression des éditions qui ne sont pas encore implémentées }

  { Etat des sorties }
  RemoveFromMenu (20224, False, stAExclure);
  { Branche crédit-bail }
  // fq 13036 - mbo - 13.01.2006 RemoveFromMenu (-20260, False, stAExclure);
  // exclure sous options état des loyers et état des dotations simulées
  RemoveFromMenu (20264, False, stAExclure);
  // BTY 02/06 FQ 13035 Dotations simulées ouvert
  // RemoveFromMenu (20266, False, stAExclure);
  { Etat comparatif des bases TP }
  RemoveFromMenu (20254, False, stAExclure);
  //{ Etiquettes }
  //RemoveFromMenu (20276, False, stAExclure);
  { Fiche Immobilisation }
  RemoveFromMenu (20278, False, stAExclure);
  { Clôture }
  // RemoveFromMenu (2416 , False, stAExclure);
  { Import/Export }
// BTY FQ 18248
//  {$IFDEF EAGLCLIENT}
//  RemoveFromMenu (20511, False, stAExclure);
//  RemoveFromMenu (20512, False, stAExclure);
//  {$ENDIF}
  // BTY 08/11/07 FQ 21703 Exclure des favoris les anciennes éditions
  RemoveFromMenu (20400, True, stAExclure);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/10/2006
Modifié le ... : 09/10/2006
Description .. : - Suppression des menus des menus suivi des règlements
Suite ........ : - Constitution de la liste des tag à supprimer pour la gestion
Suite ........ : des favoris
Mots clefs ... : 
*****************************************************************}
procedure TripotageMenuSuiviReglement ( nModule : integer; var stAExclure : string );
begin
  // Suivi clients
  if nModule=37 Then
  begin
    VH^.CCMP.LotCli := True;
    VH^.CCMP.LotFou := False;

    if VH^.PaysLocalisation=CodeISOES then
    begin
      RemoveFromMenu ( -37105, True, stAExclure);
      RemoveFromMenu ( -37205, True, stAExclure);
      RemoveFromMenu ( -37250, True, stAExclure);
      RemoveFromMenu ( -37260, True, stAExclure);
      RemoveFromMenu ( -7456, True, stAExclure);
      RemoveFromMenu ( 37502, False, stAExclure);
      RemoveFromMenu ( 37405, False, stAExclure);
      RemoveFromMenu ( 37415, False, stAExclure);
      RemoveFromMenu ( 37545, False, stAExclure);
      RemoveFromMenu ( 37560, False, stAExclure);
      RemoveFromMenu ( 37565, False, stAExclure);
    end else RemoveFromMenu ( -370100, True, stAExclure);
    //=== Pavés à oter ===
    RemoveFromMenu ( -37250, True, stAExclure); // pavé chèque
    RemoveFromMenu ( -37260, True, stAExclure); // Pavé carte bleu
    RemoveFromMenu ( -37305, True, stAExclure); // Pavé lots

    RemoveFromMenu ( 7524, False, stAExclure);
    //=== Détails des Pavés ===
    // Dans le pavé Prélèvements
    RemoveFromMenu ( 37210, False, stAExclure); // Emission de bordereau
    // Dans le pavé autres traitements
    RemoveFromMenu ( 37535, False, stAExclure);
    // autres ???
    RemoveFromMenu ( -37260, False, stAExclure);
    RemoveFromMenu ( -37305, False, stAExclure);
    RemoveFromMenu ( 37550, False, stAExclure); // saisie tréso
    // BPY le 16/12/2003 => suppression des export CFONB a la demande de la FFF
    RemoveFromMenu ( -44, False, stAExclure);
    RemoveFromMenu ( 37215, False, stAExclure);
    // fin BPY
    RemoveFromMenu ( 37600, True, stAExclure); {Lek 230106 Autres éditions}
    // SBO 05/09/2006 : code spécif pour Encaissement de masse Direct NRJ
    if not EstSpecif('51206') then     RemoveFromMenu ( 37590, False, stAExclure);
  end else if nModule = 38 then
  begin
    if not Init then
    begin
      VH^.CCMP.LotCli := False;
      VH^.CCMP.LotFou := True;
    end ;
    if VH^.PaysLocalisation=CodeISOES then
    begin
      RemoveFromMenu(38100,TRUE,stAExclure) ;
      RemoveFromMenu(-38210,TRUE,stAExclure) ;
      RemoveFromMenu(38280,TRUE,stAExclure) ;
      RemoveFromMenu(38400,TRUE,stAExclure) ;
      RemoveFromMenu(38502,FALSE,stAExclure) ;
      RemoveFromMenu(38505,FALSE,stAExclure) ;
      RemoveFromMenu(38545,FALSE,stAExclure) ;
      RemoveFromMenu(38560,FALSE,stAExclure) ;
      RemoveFromMenu(38565,FALSE,stAExclure) ;
      RemoveFromMenu(-7456,TRUE,stAExclure) ;
    end else
    begin
      RemoveFromMenu(-380100,TRUE,stAExclure) ;
    End ; //XVI 24/02/2005
    RemoveFromMenu(-38305,TRUE, stAExclure) ; // Pavé lots
    RemoveFromMenu(38210, False, stAExclure) ; // Emissions de bordereaux
    RemoveFromMenu(38290, False, stAExclure) ; // Emissions de bordereaux internationales
    RemoveFromMenu(38535, False, stAExclure) ; // Profil utilisateur
    RemoveFromMenu(38540, False, stAExclure) ; // Circuit validation décaissement
    RemoveFromMenu(38550, False, stAExclure) ; // Saisie de tréso
    // YMO 10/01/2006 FQ17118 Suppression aussi du menu "autres traitements"
    // RemoveFromMenu(38555, False, stAExclure) ; // relevé de facture
    RemoveFromMenu(7526, False, stAExclure) ; // Lettrage --> Ecart de conversion
    // BPY le 16/12/2003 => suppression des export CFONB a la demande de la FFF
    RemoveFromMenu(38125, False, stAExclure) ; // Saisie de tréso
    RemoveFromMenu(38215, False, stAExclure) ; // Saisie de tréso
    RemoveFromMenu(38297, False, stAExclure) ; // Saisie de tréso
    // SBO 05/09/2006 : code spécif pour Encaissement de masse Direct NRJ
    if not EstSpecif('51206') then RemoveFromMenu (38590, False, stAExclure);
  end ;
end;

procedure TripotageMenu (NumModule : integer);
var stAExclure : string;
begin
  if ctxPCL in V_PGI.PGIContexte
    then TripotageMenuModePCL ( [NumModule] , stAExclure )
    else TripotageMenuModePGE ( [NumModule] , stAExclure ) ;
  TripotageMenuInternat( [NumModule] , stAExclure ) ;
  if (NumModule = 20) or (NumModule = 2 ) then TripotageMenuAmortissement (stAExclure);
  if (NumModule = 37) or (NumModule = 38 ) then TripotageMenuSuiviReglement (NumModule, stAExclure);
end;

procedure AssignZoom ;
BEGIN
AssignLesZoom(False) ;
ProcGetVH:=GetVH ;
ProcGetDate:=GetDate ;
if Not Assigned(ProcZoomGene)    then ProcZoomGene    :=FicheGene;
if Not Assigned(ProcZoomTiers)   then ProcZoomTiers   :=FicheTiers;
if Not Assigned(ProcZoomSection) then ProcZoomSection :=FicheSection;
if Not Assigned(ProcZoomJal)     then ProcZoomJal     :=FicheJournal;
if Not Assigned(ProcZoomCorresp) then ProcZoomCorresp :=ZoomCorresp;
if Not Assigned(ProcZoomRub)     then ProcZoomRub     :=FicheRubrique ;

if Not Assigned(ProcZoomBudGen)  then ProcZoomBudGen  :=FicheBudgene;
if Not Assigned(ProcZoomBudSec)  then ProcZoomBudSec  :=FicheBudsect;
{$IFDEF EAGLCLIENT}
// If Not Assigned(ProcZoomNatCpt)  Then ProcZoomNatCpt  :=FicheNatCpte; // A FAIRE : Nécessite NATCPTE_TOM.PAS
{$ELSE}
  {$IFDEF V620}
  If Not Assigned(ProcZoomNatCpt)  Then ProcZoomNatCpt  :=AGLFicheNatCpte ;
  {$ELSE}
  If Not Assigned(ProcZoomNatCpt)  Then ProcZoomNatCpt  :=FicheNatCpte ;
  {$ENDIF V620}
{$ENDIF}
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 13/05/2004
Modifié le ... :   /  /
Description .. : Initialisation de la sérialisation
Mots clefs ... :
*****************************************************************}
procedure InitSerialisation;
var sDom : String ;
begin
  if V_PGI.ModePCL = '1' then
  begin
    sDom := '00280011' ;
    VH^.SerProdCompta       := '00140080';
    VH^.SerProdBudget       := '00290080';
    VH^.SerProdImmo         := '00170080';
    VH^.SerProdIcc          := '00034080';
    VH^.SerProdRevision     := '00420080';
    VH^.SerProdCre          := '00054080';
    VH^.SerProdGed          := '00119080';
    VH^.SerProdMess         := '00127080';
    VH^.SerProdAgricole     := '00446080';
    VH^.SerProdSuiviReg     := '00475080';
    VH^.SerProdGCD          := '00577080';
    VH^.SerProdDRF          := '00578080';
    VH^.SerProdRIC          := '00579080';
    VH^.SerProdDAS2         := '00580080';
    VH^.SerProdExpertETEBAC := '00609080';
    VH^.SerProdSCAN         := '00610080';

    FMenuG.SetSeria(sDom, [VH^.SerProdCompta,VH^.SerProdBudget,VH^.SerProdImmo,VH^.SerProdIcc,
                          VH^.SerProdRevision, VH^.SerProdCre, VH^.SerProdGed,
                          VH^.SerProdAgricole, VH^.SerProdMess, VH^.SerProdSuiviReg,
                          VH^.SerProdGCD,VH^.SerProdDRF,VH^.SerProdRIC,VH^.SerProdDAS2,
                          VH^.SerProdExpertETEBAC,VH^.SerProdSCAN],
                          ['Comptabilité','Budget','Immobilisation','ICC',
                          'Révision','CRE','GED',
                          'Spécificités Agricoles','Messagerie/Agenda', 'Suivi des règlements',
                          'GCD','DRF','Révision intégrée','DAS', 'Cegid Expert ETEBAC','Cegid Expert SCAN']) ;
  end else
  begin
    If CS3 Then
      BEGIN
      sDom:='00396010' ;      //BS
      VH^.SerProdCompta       := '00014080';
      VH^.SerProdBudget       := '00397080';
      VH^.SerProdImmo         := '00015080';
      FMenuG.SetSeria(sDom,[VH^.SerProdCompta,VH^.SerProdBudget,VH^.SerProdImmo],
                            ['Business Suite','Budget','Amortissement']) ;

      END Else
      BEGIN
      sDom:='05990011' ;      //BP
      VH^.SerProdCompta       := '05010080';
      VH^.SerProdBudget       := '05020080';
      VH^.SerProdImmo         := '05040080';
      VH^.SerProdEtebac       := '05030011';
      VH^.SerProdMultilingue  := '00088080'; // MODIF MULTILINGUE
      VH^.SerProdPackIFRS     := '00232080'; // MODIF IFRS
      FMenuG.SetSeria(sDom,[VH^.SerProdCompta,VH^.SerProdBudget,VH^.SerProdImmo,VH^.SerProdMultilingue,VH^.SerProdPackIFRS],
                            ['Business Place','Budget','Amortissement','Pack Traduction/Personnalisation','Pack IAS/IFRS']) ;
      END ;
  end;
end;

procedure InitApplication ;
BEGIN
    AssignZoom ;
    FMenuG.OnDispatch:=Dispatch ;
    FMenuG.OnChangeModule:=AfterChangeModule ;
    FMenuG.OnAfterProtec:=AfterProtec ;
    FMenuG.OnChargeMag:=ChargeMagHalley ;
    FMenuG.OnMajAvant:=nil;
    FMenuG.OnMajApres:=nil;
    FMenuG.OnChargeFavoris := ChargeFavoris;
    {$IFDEF EAGLCLIENT}
    V_PGI.DispatchTT:=DispatchTT ;
    AssignZoom ;
    {$ELSE}
    FMenuG.OnMajPendant:=nil;
    {$ENDIF}

    FMenuG.SetModules([6],[]) ;
    FMenuG.SetPreferences(['Saisies'],False) ;
    FMenuG.Outlook.OnChangeActiveGroup:=FMDispS5.AfterChangeGroup;
    {$IFDEF SCANGED}
    OnGlobalAfterImportGED := MyAfterImport;
    {$ENDIF}

    // GCO - 14/11/2006
    {$IFNDEF EAGLCLIENT}
    FMenuG.OnAfterUpdateBarreOutils := OnAfterUpdateBarreOutils;
    {$ENDIF}

    {$IFDEF TESTSIC}
    {$IFDEF EAGLCLIENT}
    SaveSynRegKey('eAGLHost', 'CWAS-DEV3:80', true);
    FCegidIE.HostN.Enabled := false;
    {$ENDIF EAGLCLIENT}
    {$ENDIF TESTSIC}
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 19/01/2005
Modifié le ... : 26/06/2006
Description .. : - LG - 19/01/2005 - utilisation de la fct FicheTiersZoom. On
Suite ........ : recup le range pour ouvrir la fiche des tiers sur l'auxi courant
Suite ........ : - LG - 26/06/2006 - FB 18437 - on remet ajour al liste des
Suite ........ : comptes apres l'appel des zoom
Mots clefs ... :
*****************************************************************}
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
  case Num of
    // general
    1,7112 : begin
              FicheGene(Nil,'',LeQuel,Action,0);
              if GetFolioCourant <> nil then
               GetFolioCourant.EffacerListeMemoire ;
             end ;
    // tiers
    2,7145 : begin
              FicheTiersZoom(Nil,'',LeQuel,Action,1,Range) ;
              if GetFolioCourant <> nil then // FQ 18711 : SBO 31/08/2006
              GetFolioCourant.EffacerListeMemoire ;
             end ;
    // journal
    4,7211 : begin
              FicheJournal(Nil,'',Lequel,Action,0) ;
              if GetFolioCourant <> nil then // FQ 18711 : SBO 31/08/2006
              GetFolioCourant.EffacerListeMemoire ;
             end ;

    // tables libres
    5 : begin
        {$IFDEF V620}
          if (not (pos('GENE',TT) = 0)) then AGLFicheNatCpte(nil,'G0'+AnsiLastChar(TT),'',Action,1)
          else if (not (pos('SECT',TT) = 0)) then AGLFicheNatCpte(nil,'S0'+AnsiLastChar(TT),'',Action,1)
          else if (not (pos('TIERS',TT) = 0)) then AGLFicheNatCpte(nil,'T0'+AnsiLastChar(TT),'',Action,1)
          else AGLFicheNatCpte(nil,'I0'+AnsiLastChar(TT),'',Action,1);
        {$ELSE}
          if (not (pos('GENE',TT) = 0)) then FicheNatCpte(nil,'G0'+AnsiLastChar(TT),'',Action,1)
          else if (not (pos('SECT',TT) = 0)) then FicheNatCpte(nil,'S0'+AnsiLastChar(TT),'',Action,1)
          else if (not (pos('TIERS',TT) = 0)) then FicheNatCpte(nil,'T0'+AnsiLastChar(TT),'',Action,1)
          else FicheNatCpte(nil,'I0'+AnsiLastChar(TT),'',Action,1);
        {$ENDIF V620}
        end;

    // ???
    900 : ParamRegroupementMultiDossier (Lequel); //AGLLanceFiche('Y','MULTIDOSSIER','',Lequel,ActionToString(Action)) ;   {YYMULTIDOSSIER}
    (* 17/08/2005 - CA - FQ 15713 et pour conformité avec les autres applications PGI
    995 : ParamTable(TT,Action,0,Nil,9) ;   {choixext}
    997 : ParamTable(TT,Action,0,Nil,6) ;   {choixext}
    999 : ParamTable(TT,Action,0,Nil,3) ;   {choixcode ou commun}                     *)

    // pour mise à jour par defaut des tablettes
    994 : ParamTable(TT,taModif,0,Nil,9) ;   {choixext}
    995 : ParamTable(TT,taCreat,0,Nil,9) ;   {choixext}
    996 : ParamTable(TT,taModif,0,Nil,6) ;   {choixext}
    997 : ParamTable(TT,taCreat,0,Nil,6) ;   {choixext}
    998 : ParamTable(TT,taModif,0,Nil,3) ;   {choixcode ou commun}
    999 : ParamTable(TT,taCreat,0,Nil,3) ;   {choixcode ou commun}
    {JP 17/05/06 : FQ 18072 : accès à la fiche des utilisateurs}
    3170 : AglLanceFiche('YY','YYUTILISAT', Range, Lequel, ActionToString(Action));

    7701 : ParametrageRubrique(Lequel,Action, CtxRubrique) ;
    // Comptes de correspondance
    1325 : CCLanceFiche_Correspondance('GE');	// Plan de correspondance Généraux
    1330 : CCLanceFiche_Correspondance('AU');	// Plan de correspondance Auxiliaire

    // Section
    71781 : FicheSection(nil,'A1',Lequel,Action,1);
    71782 : FicheSection(nil,'A2',Lequel,Action,1);
    71783 : FicheSection(nil,'A3',Lequel,Action,1);
    71784 : FicheSection(nil,'A4',Lequel,Action,1);
    71785 : FicheSection(nil,'A5',Lequel,Action,1);

    // compteur
    72111 : begin
              case Action of
                taCreat : YYLanceFiche_Souche('CPT','','ACTION=CREATION;CPT');
                taModif :
                        { FQ 19072 BVE 12.04.07 }
                        begin
                          if Lequel <> '' then
                             YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT')
                          else
                             YYLanceFiche_Souche('CPT','','ACTION=CREATION;CPT');
                        end;
                        { END FQ 19072 }
                else      YYLanceFiche_Souche('CPT','','ACTION=CONSULTATION;CPT');
              end;
            end;
    {$IFDEF AMORTISSEMENT}
    // Lieu géographiques
    2342 : ParamTable('TILIEUGEO',taCreat,2342000,nil) ;
    // Motifs de sortie
    2345 : ParamTable('TIMOTIFCESSION',taCreat,2345000,nil) ;
    2348 : ParamTable('AMREGROUPEMENT',taCreat,0,nil) ;
    {$ENDIF}

    17141 : CpLanceFiche_TypeVisaFiche(Range, Lequel, '');
    //17142 : ParamTable('CPCIRCUITBAP', taCreat, 150, nil);
    //17143 : ParamTable('CPTYPEMAILBAP', taCreat, 150, nil);
    17144 : CpLanceFiche_CircuitFiche(Lequel, iif(Action = taCreat, 'C', iif(Action = taModif, 'M', 'L')));
    end ;
end ;

procedure TFMDispS5.FMDispS5Create(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  PGIAppAlone:=True ;
  CreatePGIApp ;
{$ENDIF}
end;

procedure TFMDispS5.AfterChangeGroup(Sender : TObject);
var
  n : Integer;
begin
  if ctxPCL in V_PGI.PGIContexte then
  begin
    if (NumModule=55) then
    begin
      { FQ 15736 - Si dernier groupe de la liste, alors c'est le budget ==> popup particulier }
      if (fMenuG.Outlook.Groups.Count = fMenuG.Outlook.ActiveGroup+1) then begin
        ChargeMenuPop(integer(hm3),FMenuG.DispatchX);
        {JP 24/04/07 : FQ 15945 : si pas d'analytique, on désactive les popup analytique.
                       Impossible de les rendre invisible, l'agl intervenant après}
        if GetParamSoc('SO_CPPCLSANSANA') then begin
          for n := 0 to PopHalley.Items.Count - 1 do begin
            if (PopHalley.Items[n].Tag = 26165) or
               (PopHalley.Items[n].Tag = 26185) then PopHalley.Items[n].Enabled := False;
          end;
        end;
      end
      else ChargeMenuPop(integer(hm2),FMenuG.DispatchX);
    end
    else if (assigned (FMenuG.Outlook)) and (FMenuG.Outlook.ActiveGroup<>0) then ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/11/2006
Modifié le ... : 13/11/2006
Description .. : Création ou suppression du Bouton de Statut Fiscal du DOSSIER
Mots clefs ... :
*****************************************************************}
procedure TFMDispS5.InitBStatutLiasse;
begin
  if VH^.BStatutLiasse <> nil then Exit;
  FMDispS5.BStatutLiasse := FMenuG.SetBouton(60, '', 'Statut du dossier (liasse fiscale)', 54740, 3, FMDispS5.ImageList);
  FMDispS5.BStatutLiasse.Name := 'BSTATUTLIASSE';
  FMDispS5.BStatutLiasse.Images   := FMDispS5.ImageList;
  FMDispS5.BStatutLiasse.ShowHint := True;
  FMDispS5.BStatutLiasse.OnClick  := FMDispS5.OnClickBStatutLiasse;
  FMDispS5.BStatutLiasse.ImageIndex := GetParamSocSecur('SO_CPSTATUTDOSSIERLIASSE', 3);
  VH^.BStatutLiasse := FMDispS5.BStatutLiasse;
  FMDispS5.BStatutLiasse.Visible := GetParamSocSecur('SO_CPCONTROLELIASSE', '') <> '';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMDispS5.OnClickBStatutLiasse( Sender : TObject);
begin
  ControleStatutLiasse;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/01/2007
Modifié le ... : __/__/____
Description .. : Création ou suppression du Bouton de Statut de révision
Mots clefs ... :
*****************************************************************}
procedure TFMDispS5.InitBStatutRevision;
begin
  if VH^.BStatutRevision = nil then
  begin
    FMDispS5.BStatutRevision            := FMenuG.SetBouton(85, '', 'Statut du dossier (révision)', 54750, 3, FMDispS5.ImageList);
    FMDispS5.BStatutRevision.Name       := 'BSTATUTREVISION';
    FMDispS5.BStatutRevision.Images     := FMDispS5.IMREVISION;
    FMDispS5.BStatutRevision.ShowHint   := True;
    FMDispS5.BStatutRevision.OnClick    := FMDispS5.OnClickBStatutRevision;
    FMDispS5.BStatutRevision.ImageIndex := GetParamSocSecur('SO_CPSTATUTREVISION', 3);
    VH^.BStatutRevision                 := FMDispS5.BStatutRevision;
  end;  
  AccesBStatutRevision;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/01/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFMDispS5.OnClickBStatutRevision( Sender : TObject );
begin
  ControleStatutRevision;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFMDispS5.ExercClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  // CA - 27/11/2001
  ParamSociete(False,'','SCO_DATESDIVERS','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,0) ;
{$ENDIF}
end;

Procedure InitLaVariablePGI;
Begin
  {Version}
  Apalatys:='CEGID' ;
  HalSocIni:='CEGIDPGI.INI' ;
  Copyright:='© Copyright ' + Apalatys ;
  V_PGI.CodeProduit  := '001';
  If CS3 Then RenseignelaSerie(ExeCCS3) Else RenseignelaSerie(ExeCCS5) ;

  {Généralités}
  V_PGI.VersionDemo:=True ;
  V_PGI.SAV:=False ;
  V_PGI.VersionReseau:=True ;
  V_PGI.PGIContexte:=[ctxCompta] ;
  V_PGI.CegidAPalatys:=FALSE ;
  V_PGI.CegidBureau:=TRUE ;
  V_PGI.StandardSurDP:=True ;
  V_PGI.MajPredefini:=False ;
  V_PGI.MultiUserLogin:=False ;
  V_PGI.BlockMAJStruct:=True ;
  V_PGI.CegidUpdateServerParams:= 'www.update.cegid.fr';
{$IFDEF EAGLCLIENT}
  V_PGI.AutoSearch := False;
{$ELSE}
{$ENDIF}
  V_PGI.EuroCertifiee:=False;

  {Série}
  If CS3 Then V_PGI.LaSerie:=S3 Else V_PGI.LaSerie:=S5 ;
  ChargeXuelib ;
  V_PGI.PortailWeb := 'http://clients.cegid.fr';
  
  V_PGI.NumMenuPop:=27 ;
  //  V_PGI.OutLook:=TRUE ; // MAJ par l'AGL, ne pas toucher !!
  V_PGI.OfficeMsg:=True ;
  V_PGI.NoModuleButtons:=FALSE ;
  V_PGI.NbColModuleButtons:=1 ;
  {Divers}
  V_PGI.MenuCourant:=0 ;
  V_PGI.OKOuvert:=FALSE ;
  V_PGI.Halley:=TRUE ;
{$IFDEF EAGLCLIENT}
  V_PGI.VersionDEV:=false ;
{$ELSE}
  V_PGI.VersionDEV:=TRUE ;
{$ENDIF}
  V_PGI.ImpMatrix := True ;
  V_PGI.DispatchTT:=DispatchTT ;
  V_PGI.ParamSocLast:=False ;
  //V_PGI.Decla100:=TRUE ;
  V_PGI.RAZForme:=TRUE ;
  //SG6 07.03.05 FQ 15389
  //  V_PGI.SansMessagerie := True;
  V_PGI.TabletteHierarchiques := True;
  If CS3 Then RenseignelaSerie(ExeCCS3) Else RenseignelaSerie(ExeCCS5) ;
  // Initialisé à Vrai par défaut !
  V_PGI.EnableTableToView := False ;
End;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 25/04/2006
Modifié le ... : 25/04/2006
Description .. : Reprise de la fonction de la ComSx pour pouvoir lancer la
Suite ........ : compta en ligne de commande avec les paramètres :
Suite ........ : /USER=xxx /PASSWORD=xxx /DOSSIER=xxx
Suite ........ : /DATE=xxx /RECALCUL
Suite ........ : 
Suite ........ : Lance le recalcul des soldes de tous les comptes
Suite ........ : 
Suite ........ : !!! Attention !!! le dpr doit être modiofié de la sorte :
Suite ........ : var
Suite ........ :   i              : integer;
Suite ........ :   St           : string;
Suite ........ :   OkArg    : Boolean;
Suite ........ :   OkLogin : Boolean;
Suite ........ : begin
Suite ........ : 
Suite ........ :   Application.Initialize;
Suite ........ :   InitAGL;
Suite ........ :   Application.Title := 'Comptabilité S5';
Suite ........ :   Application.HelpFile := 'CCS5.HLP';
Suite ........ :   Application.CreateForm(TFMenuG, FMenuG);
Suite ........ :   OkArg   := FALSE ;
Suite ........ :   OkLogin := FALSE ;
Suite ........ :   for i:=1 to ParamCount do
Suite ........ :     begin
Suite ........ :     St := ParamStr(i);
Suite ........ :     if (pos('/USER', St) <> 0) then
Suite ........ :       OkLogin := TRUE;
Suite ........ :     if (pos('/RECALCUL', St) <> 0) then
Suite ........ :       OkArg := TRUE;
Suite ........ :     end;
Suite ........ :   if OkArg  and OkLogin then
Suite ........ :      LanceMenuAuto
Suite ........ :   else
Suite ........ :     begin
Suite ........ :     Application.CreateForm(TFMDispS5, FMDispS5);
Suite ........ :     InitApplication ;
Suite ........ :     end;
Suite ........ : 
Suite ........ :   SplashScreen:=Nil ;
Suite ........ :   if ParamCount=0 then
Suite ........ :      BEGIN
Suite ........ :      SplashScreen:=TSplashScreen.Create(Application) ;
Suite ........ :      SplashScreen.Show ;
Suite ........ :      SplashScreen.Update ;
Suite ........ :      END ;
Suite ........ : 
Suite ........ :   Application.Run ;
Suite ........ :  end.
Mots clefs ... :
*****************************************************************}
procedure LanceMenuAuto;
var Connect,B1,B2      : Boolean;
    i,j                : integer;
    Nom, St,Value      : string;
    Num                : integer;
    P2                 : THPanel ;
    vBoRecalcul        : Boolean ;
begin
  V_PGI.Debug       := FALSE ;
  V_PGI.Versiondev  := FALSE ;
  V_PGI.Synap       := FALSE ;
  VH^.GrpMontantMin := 0 ;
  VH^.GrpMontantMax := 1000000 ;
  V_PGI.DateEntree  := Date ;
  VH^.Mugler        := FALSE ;
  V_PGI.Halley      := TRUE ;
  V_PGI.NumVersion  := '3' ;
  V_PGI.SAV         := TRUE ;
  V_PGI.Versiondev  := FALSE ;
  V_PGI.Synap       := FALSE ;
  VH^.GereSousPlan  := True ;
  V_PGI.Halley      := TRUE ;
  VH^.ModeSilencieux := True;
  V_PGI.MultiUserLogin := TRUE;
  Num               := 0;

  if ((V_PGI.ModePCL='1')) then
    begin
    V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
    V_PGI.SAV:=FALSE ;
    end;

  Connect := FALSE;
  vBoRecalcul := False ;
  P2      := nil;
  for i:=1 to ParamCount do
    begin
    St    := ParamStr(i) ;
    Nom   := UpperCase(Trim(ReadTokenPipe(St,'='))) ;
    Value := UpperCase(Trim(St)) ;

    //Paramètres de connexion
    if Nom='/USER'     then   V_PGI.UserLogin  := Value
    else if Nom='/PASSWORD' then   V_PGI.PassWord   := Value
    else if Nom='/DATE'     then   V_PGI.DateEntree := StrToDate(Value)
    else if Nom='/TAG'      then   Num := VALEURI(Value)
    else if Nom='/DOSSIER'  then
      BEGIN
      if V_PGI.PassWord = '000' then
        V_PGI.PassWord := CryptageSt(DayPass(Date));
      V_PGI.CurrentAlias := Value;
      V_PGI.DefaultSectionName := '';
      j := pos('@', Value);
      if (j > 0) then
        begin
        V_PGI.DefaultSectionName := Copy(Value, j + 1, 255);
        V_PGI.CurrentAlias := Copy(Value, 1, j - 1);
        if (j > 3) and (Copy(Value, 1, 2) = 'DB') then
          V_PGI.NoDossier := Copy(Value, 3, j - 3); // sinon reste à '000000'
        end;
      V_PGI.RunFromLanceur := (V_PGI.DefaultSectionName <> '');
{$IFNDEF EAGLCLIENT}
      Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL,NIL);
{$ELSE}
      Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL);
{$ENDIF}
      if not Connect then break;
      if v_pgi.DefaultSectionName <> '' then
        if (v_pgi.DBName = v_pgi.DefaultSectionDBName) then
          v_pgi.InBaseCommune := true;
      END
    else if (Nom='/RECALCUL') then vBoRecalcul := True ;
  end; //FOR

  if (not Connect) then
  begin
  if V_PGI.PassWord = '000' then
    V_PGI.PassWord := CryptageSt(DayPass(Date));
{$IFNDEF EAGLCLIENT}
  Connect := ConnecteHalley(V_PGI.CurrentAlias,TRUE,@ChargeMagHalley,NIL,NIL,NIL);
{$ELSE}
  Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL);
{$ENDIF}
  end;

  if vBoRecalcul then
    ModeRecalcul
  else if Num > 0 then
    begin
    B1  := false;
    B2  := false;
    Dispatch (Num, P2, B1, B2);
    end ;

  Application.ProcessMessages;
  if Connect then
    begin
{$IFNDEF EAGLCLIENT}
    Logout ;
{$ENDIF}
    DeconnecteHalley ;
    end ;

  // Fermeture de l'application
  Application.ProcessMessages;
  if FMenuG <> nil then
    begin
    FMenuG.ForceClose := True ;
    PostMessage(FmenuG.Handle,WM_CLOSE,0,0) ;
    FMenuG.Close ;
    end;
{$IFNDEF EAGLCLIENT}
  If (DBSOC<>NIL) AND (DBSOC.Connected) then DBSOC.Connected:=FALSE ;
{$ENDIF}
  VH^.STOPRSP:=TRUE ;
  SourisNormale ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 25/04/2006
Modifié le ... : 25/04/2006
Description .. : Lance le recalcul des soldes de tous les comptes
Mots clefs ... :
*****************************************************************}
function ModeRecalcul : boolean ;
var i  : integer ;
    St : string ;
begin

  result := False ;

  for i:=1 to ParamCount do
    begin
    St := ParamStr(i) ;

    if UpperCase(Trim(ReadTokenPipe(St,'='))) ='/RECALCUL' then
      begin
      //Lancement du recalcul
      MajTotTousComptes(False, '');
      result := True ;
      break ;
      end ;

    end ;

end ;

Procedure EnvoiBL;
{$IFDEF NETEXPERT}
var
    NoSEqNet,reponse    : integer;
    Fileini,TypeExport  : string;
    OutTOBNet           : TOB;
  {$ENDIF}
begin
{$IFDEF NETEXPERT} // lien Netexpert
          if (IsDossierNetExpert(V_PGI.NoDossier, NoSEqNet)) and
          ((ExisteSQL('SELECT E_IO FROM ECRITURE WHERE E_IO="X" or E_PAQUETREVISION=1'))
             or ExisteSQL ('SELECT CHT_NOMTABLE FROM CPHISTOPARAM ')) then
          begin
                 if GetParamSocSecur ('SO_CPSYNCHROAUTOMATIQUE', TRUE) then Reponse := mryes
                 else
                  reponse := PGIAsk('Vous avez effectué des modifications sur votre dossier, voulez-vous synchroniser avec '+RechDom('CPLIENCOMPTABILITE',VH^.CPLienGamme, False) +' ?','Envoyer');
                 if reponse=mrYes then
                 begin
                     OutTOBNet := InterrogeCorbeil ('COMPTA', 'TRA', FALSE);
                     if (OutTOBNet <> nil) and (OutTobNet.detail.count > 0)  then
                     begin
                           PGIInfo('Synchronisation avec Business Line impossible.'+
                          #10#13+' Il reste des fichiers en attente d''intégration.'+
                          #10#13+' Vous devez les intégrer avant de pouvoir effectuer un envoi ');

                     end
                     else
                     begin
                        Fileini := GetEnvVar('TEMP') + '\COMSX.ini';
                        if NoSEqNet >= 1 then TypeExport := 'SYN'
                        else TypeExport := 'DOS';
                        GenereIniExport(Fileini, 'PG'+V_PGI.NoDossier+'_'+IntTostr(NoSEqNet)+'.TRA', TypeExport);
                        ExportDonnees(Fileini+';EXPORT;Minimized',TRUE) ;
                     end;
               end;
          end;
{$ENDIF}
end;

Procedure ControlEnvoiBL;
{$IFDEF NETEXPERT} // lien Netexpert
var                                         
DateSynchro : TDateTime;
{$ENDIF}
begin
{$IFDEF NETEXPERT} // lien Netexpert
        DateSynchro := iDate1900;
        if (GetParamSocSecur('SO_CPDATESYNCHRO2', '') <> 0) and (GetParamSocSecur ('SO_CPDATESYNCHRO2', '') <> iDate1900) then DateSynchro := GetParamSocSecur ('SO_CPDATESYNCHRO2', '')
        else   if (GetParamSocSecur('SO_CPDATESYNCHRO1', '') <> 0) and (GetParamSocSecur ('SO_CPDATESYNCHRO1', '') <> iDate1900) then DateSynchro := GetParamSocSecur ('SO_CPDATESYNCHRO1', '');
        if DateSynchro >iDate1900 then
        begin
             DateSynchro := DateSynchro + 14;
        end;
        if (DateSynchro <> iDate1900) and (now > DateSynchro) then
        begin
                  if (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE) = TRUE) and
                  ((ExisteSQL('SELECT E_IO FROM ECRITURE WHERE E_IO="X" or E_PAQUETREVISION=1'))
                     or ExisteSQL ('SELECT CHT_NOMTABLE FROM CPHISTOPARAM ') )then
                     PGIInfo('Vous n''avez pas effectué d''envoi ou de réception de fichier de synchronisation depuis 14 jours. '+
                    #10#13+' Nous vous rappelons qu''une bonne synchronisation de données nécessite des échanges réguliers');
        end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2006
Modifié le ... : 25/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure OnAfterUpdateBarreOutils(Sender : TObject);
begin
  if (TToolBarButton97(Sender).Name = 'BSTATUTLIASSE') then
  begin
    if VH^.BStatutLiasse <> nil then
      VH^.BStatutLiasse.Visible := GetParamSocSecur('SO_CPCONTROLELIASSE', '') <> '';
  end;

  if (TToolBarButton97(Sender).Name = 'BSTATUTREVISION') then
  begin
    if VH^.BStatutRevision <> nil then
      AccesBStatutRevision;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
CS3:=FALSE ;
If Pos('CCS3',ParamStr(0))>0 Then CS3:=TRUE ;
ProcChargeV_PGI :=  InitLaVariablePGI ;
RegisterHalleyWindow ;
end.

















