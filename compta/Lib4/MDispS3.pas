{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
unit MDispS3 ;

interface

uses
{$IFDEF EAGLCLIENT}
    utileagl,
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
    eSaisieTr,
    CPMODIFECHEMP_TOF,    
{$ELSE}
    MenuOLG,
    uedtcomp,
    Tablette,
    FE_Main,
    EdtEtat,
    YNewMessage,  // pour ShowNewMessage
{$ENDIF EAGLCLIENT}
    Forms,
    sysutils,
    HMsgBox,
    Classes,
    Controls,
    HPanel,
    Hctrls,
    UIUtil,
    ImgList,
    AglInitPlus,
    AGLInit,
    uMultiDossier,
{$IFDEF RR}
    CPANOECR,
{$ENDIF}
    HEnt1,        // TActionFiche
    GalOutil,
    AglMail       // pour AglMailForm
    {$IFDEF NETEXPERT}
    ,UNetListe
    ,UAssistComsx
    ,uNEActions
    {$ENDIF}
    ,reseau
    {$IFDEF EDICOMPTA}
    ,Echg_Code
    {$ENDIF}
    ,CPGestionAna_Tof //CPLanceGestionAna //SG6 08/12/2004
    ;



{$IFNDEF EAGLCLIENT}
type
  TFMDispS3 = class(TDatamodule)
    procedure FMDispS3Create(Sender: TObject);
    private
      NumModule : integer;
      procedure ExercClick(Sender: TObject);
    public
      SeriaMessOk : Boolean;
      procedure AfterChangeGroup ( Sender : TObject) ;
    end ;

Var FMDispS3 : TFMDispS3 ;
{$ENDIF}

{$IFDEF SCANGED}
procedure MyAfterImport (Sender: TObject; FileID: Integer; var Cancel: Boolean) ;
{$ENDIF}

{$IFDEF EAGLCLIENT}
 Function  OkRevis : boolean ;
{$ELSE}
function MyAglMailForm ( var Sujet, A, CC: String ; Body: TStringList ; var Files: String ;
    ZoneTechDataVisible: boolean=True) : boolean ;
Function  OkRevis : boolean ;
procedure ControleLettrageEnSaisie ;
{$ENDIF}
function  AffichageEcranAccueil : boolean;

Function  GetStRevis : String ;
Procedure DispatchPCL10 ;
procedure LanceEtatLibreS5 ;
procedure LanceEtatLibreS7 ( NatureEtat : String ) ;
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
procedure CacherMenuPackAvance ( const TNumModule : array of integer; var stAExclure : string ); // Remplace VireMenuPackAvance
procedure TripotageMenuModePGE ( const TNumModule : array of integer; var stAExclure : string );
procedure TripotageMenu (NumModule : integer);
procedure AssignZoom ;
procedure InitApplication ;
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
Procedure InitLaVariablePGI;
procedure InitSerialisation;

//************************************************************************************************//
//************************************************************************************************//
//************************************************************************************************//
implementation

uses
    Windows,
    Messages,
    Menus,
    Audit,
  {$IFDEF TRSYNCHRO}
    Commun, {Pour l'initialisation des agences, menu 3217}
    UTomBanque, {Fiche banque}
    TomAgence, {Agences bancaires}
    TRMULCOMPTEBQ_TOF, {Liste des comptes bancaires}
  {$ENDIF TRSYNCHRO}
    CPMulRub_TOF,
    Rubrique_TOM,
    CPQRTRUBRIQUE_TOF,
    CPSUPPRUB_TOF,
    CPCONTROLERUB_TOF,
    CPTOTRUB_TOF,
    ESPCPQRIRPF115_TOF,  // TVA Espagnole
{$IFDEF EAGLCLIENT}
    CEGIDIEU, Hcompte,
    uTOFCPMulGuide, CPMULLETTR, TofBalagee, UTOFMULRELCPT, UTOFMULRELFAC,
    AnulCloSERVEUR,
    CPTVAMODIF_TOF, CPTVAACC_TOF, CPQRTVAFAC_TOF, CPVALTVA_TOF,
//    CPAXE_TOM,
     PAYS_TOM, CPREGION_TOF, CPCODEPOSTAL_TOF, CPTVATPF_TOF, CPSTRUCTURE_TOF,
    SUIVCPTA_TOM, CPCODEAFB_TOF, CPGENERAUX_TOM, CPTIERS_TOM, CPSECTION_TOM, CPJOURNAL_TOM,
    uTofConsEcr, CPMULANA_TOF,
    DEVISE_TOM,
    REFAUTO_TOM,
    VALSIMU_TOF, PLANGENE_TOF, PLANAUXI_TOF, PLANSECT_TOF, PLANJOUR_TOF,
    REPARCLE_TOM, {JP 08/06/05 : migration CWAS}
    CPVIDESECTIONS_TOF, {JP 10/06/05 : migration CWAS}
    CPMODIFTABLELIBRE_TOF, {JP 10/06/05 : migration CWAS}
    CPCUMULPERIODIQUE_TOF, {JP 24/06/05 : migration CWAS}

{$IFDEF PORTAGEBUDGET}
    MulMvtBu_TOF, eSaisBud, SupEcrBU_TOF, MULBUDG_TOF, BUDGENE_TOM, MULBUDS_TOF, BUDSECT_TOM,
    MULBUDJ_TOF, SUPPRBUDG_TOF, SUPPRBUDS_TOF, SUPPRBUDJ_TOF, PLANBUDGET_TOF, PLANBUDSEC_TOF,
    PLANBUDJAL_TOF, REPARTBU_TOM, CroisCpt, TTCATJALBUD_TOT, CreRubBu, CtrBud, GenToBud, CodBuRup,
    CreCodBu, CreSecBu, Csebucat, MULVALBU_TOF, COPIEBUD_TOF, RevBuRea, COPIEBUDMU_TOF, OUVREFERME_TOF,
    BALBUDTEGEN_TOF, BALBUDTESEC_TOF, BALBUDTEGES_TOF, BALBUDTESEG_TOF, BROBUDGEN_TOF, BALBUDECGEN_TOF,
    BALBUDECSEC_TOF, BALBUDECGES_TOF, BALBUDECSEG_TOF,
{$ENDIF}

{$ELSE}
    VideSect,
    ParamLib,
    EdtREtat,    // LanceEtat
    DBTables,    // TQuery
    QR,
    PGIExec,
    ConsEcr,
    SelGuide, MULLETTR, RelCompt, RelevTie, AnulClo,
    CPTVATPF_TOF,
    CONTABON_TOM,
    TvaModif, TvaAcc, QRTvaFac, QRTvaAcE, TvaValid,
    OuvrFerm,
//    Axe,
    Pays, Region, CodePost,
    DEVISE_TOM,
    CPGeneraux_TOM,
    CPTiers_TOM,
    CPJournal_TOM,
    CPSection_TOM,
    Structur, Scenario, CODEAFB,UlibEcriture ,
    CTableauBord, UGedFiles, Extourne, Constant,
    Guide, MulBudg, Mulbuds, Mulbudj, MulMvtBu, Traduc, EdtDoc, SAISUTIL,
    PlanRef, RUPTURE, REPARCLE, EtapeReg, SuivDec, StrucBud, GenToBud, CodBuRup,
    CreCodBu, CreSecBu, Csebucat, repartbu, EtbMce, CreerSoc, InitSoc,  AssistSO, CopySoc,
    MulLog, CtrlFic, CtrlDiv, ReIndex, TofRubImport,
    //calceuro,
    UTOFEnregStandard,
    SaisieTr, ModTlEcr, VisuEtebac, QRBrouil, ValidSim,
    MulAbo, MulAna, QrBroAna, QRCumulG, QRJCpte, QRGLANA,
    QRGLGESE, QRGLSEGE, QRBlAna, QRBaGeAu, QRBaAuGe, QRBLGESE, QRBLSEGE, Edtlegal, QRBRT, QRBAGTP, QRGLATP, QRBAVTP, QRGLVTP,
    EncaDeca, EcheModf, QRJuSold, QREchean, TofBalagee, QRGLAge, QRGLVen, QRGLAuxL,
    CFONBAT, CoutTier, ModEcheMP, SuivTres, CtrlCai, TVAEdite, Revision, QRLegal, MulBds,
    CPAFAUXCPTA_TOF, QRJDivis, QRBalGen, CtrBud, SaisBud, Supecrbu, MulValBu, CopieBud,
    CpyBudMu, RevBuRea, QrBroBuG, QrBroBuS, QRBLteG, QRBLteS, QRBLteGS, QRBLteSG, QRBLecG, QRBLecS, QRBLecGS, QRBLecSG,
    //Gestion croisaxe SG6 22.02.05
    ImputAna, CPREIMPUTANA_TOF,
    PssToTli, // pas encore compatible eAGL
    CopiTabL, // pas encore compatible eAGL
    QRCatSG, CreRubBu, CroisCpt, CtrRubBu, QRPLANGE, QRTier, QRSectio, QRPlanJo, Suprbudg, Suprbuds, Suprbudj,
    QRBudGen, QRBudSec, QRBudJal,
    QRBalAux, QRGLGen, QRGLAux, QRPointg, QRRappro,
    Qrjustbq, QRGLAuGe, QRGLGeAu, QrBalAge, QrBalVen, TOFEMPIMP,
    Budgene, Budsect,
  {$IFDEF CERIC}
    FonctionCERIC,            // Menu Spécifique CERIC
  {$ENDIF CERIC}
  {$IFDEF SCANGED}
    YNewDocument,
  {$ENDIF SCANGED}
{$ENDIF EAGLCLIENT}

// FICHIERS COMMUNS
    CPQRTVACTRL_TOF, // YMO 25/01/2005 Branchement Etat de contrôle en 2 tiers
    Ent1, CALCOLE,
    Choix,       // Choisir
    UTOFMULPARAMGEN, UtilSoc, MULDLETT, RELANCE_TOM,
    RappAuto, BonPayer, UTOFCPJALECR, UTOFCPGLGENE, UTOFCPGLAUXI, CPGRANDLIVRETL_TOF, CPBALGEN_TOF, CPBALAUXI_TOF,
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
    LibChpLi,     // FQ 16256 SBO 01/08/2005
    // fa 2005.08.31
    // CPMULSUPPRANA_TOF, // SBO 26/08/2005 : remplacement fiche 2 tiers par fiche cwas // Axe
    CPAXE_TOM,         // SBO 26/08/2005 : remplacement fiche 2 tiers par fiche cwas // supprana
    Saisie
{$IFDEF AMORTISSEMENT}
    , AMLISTE_TOF
    , AMEDITION_TOF
    , AMANALYSE_TOF
    , AMINTEG_TOF
    , IMMOCPTE_TOM
    , AMCLOTURE_TOF
  {$IFDEF EAGLCLIENT}
  {$ELSE}
    , ImEdGene
  {$ENDIF}
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
    {$IFNDEF EAGLCLIENT}
    , ImpCegid
    , AmExport
    {$ENDIF}
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
    , GenerAbo            // Abonnements
  {$IFDEF EAGLCLIENT}
    , CPMulABo_TOF
    , CPReconAbo_TOF
    , Contabon_TOM
  {$ELSE}
    //, ContaBon
    , ReconAbo
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
    {$IFDEF MULTISOC}
    ,YYBUNDLE_TOF
    {$ENDIF MULTISOC}
    ,CLIENSETAB_TOM
    , UtofClotureDef
    ;


{$R *.DFM}

//************************************************************************************************//
//************************************************************************************************//
//************************************************************************************************//
{$IFDEF SCANGED}
procedure MyAfterImport (Sender: TObject; FileID: Integer; var Cancel: Boolean) ;
// Evenement après scannérisation : retourne le FileID du fichier scanné
var ParGed : TParamGedDoc;
begin

  if FileId = -1 then exit;

  // ---- Insertion dans la GED sans interaction utilisateur ----
  if TForm(Sender) is TFSaisieFolio then
    SaisBorMyAfterImport (Sender,FileID,Cancel)
  else  if TForm(Sender) is TFSaisie then
    SaisieMyAfterImport (Sender,FileID,Cancel)
  else
    // ---- Insertion classique dans la GED avec boite de dialogue ----
    begin
    // Propose le rangement dans la GED
    ParGed.DocId := 0;
    ParGed.NoDossier := V_PGI.NoDossier;
    ParGed.CodeGed := '';
    // FileId est le n° de fichier obtenu par la GED suite à l'insertion
    ParGed.FileId := FileId;

    // Description par défaut du document à archiver...
    if Sender is TForm then ParGed.DefName := TForm(Sender).Caption;

    Application.BringToFront;

    if ShowNewDocument(ParGed)='##NONE##' then
      // Fichier refusé, suppression dans la GED
      V_GedFiles.Erase(FileId);
    end;
end;
{$ENDIF}

//   XXXX  XXXX  XXXX  X
//   X     X  X  X     X
//   XXXX  XXXX  X XX  X
//   X     X  X  X  X  X
//   XXXX  X  X  XXXX  XXXX
{$IFDEF EAGLCLIENT}

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Indique si les modules de révision sont accessibles.
Suite ........ : MODIF PACK AVANCE
Mots clefs ... : PACK AVANCE;REVISION
*****************************************************************}
Function OkRevis : boolean ;
BEGIN
  Result := V_PGI.Controleur ;
  if Not Result then
    PGIBox('Fonction interdite : vous devez être réviseur pour accéder aux fonctions de révision','Comptabilité S5 Pack avancé') ;
END ;


//  X   X  XXXX  X   X    XXXX  XXXX  XXXX  X
//  XXX X  X  X  XXX X    X     X  X  X     X
//  XXX X  X  X  XXX X    XXXX  XXXX  X XX  X
//  X XXX  X  X  X XXX    X     X  X  X  X  X
//  X   X  XXXX  X   X    XXXX  X  X  XXXX  XXXX
{$ELSE}

function MyAglMailForm ( var Sujet, A, CC: String ; Body: TStringList ; var Files: String ;
    ZoneTechDataVisible: boolean=True) : boolean ;
begin;
    // Routage vers la messagerie PGI au lieu du send mail agl
  Result := ShowNewMessage(0, 0, '', A);
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
BEGIN
  Result := V_PGI.Controleur ;
  if Not Result then
    PGIBox('Fonction interdite : vous devez être réviseur pour accéder aux fonctions de révision','Comptabilité S5 Pack avancé') ;
END ;

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
{$ENDIF}

function AffichageEcranAccueil : boolean;
var stAccueil : string;
    bRetour : boolean;
begin
  bRetour := False;
  stAccueil := GetParamSoc ('SO_CPECRANACCUEIL');
  if stAccueil = 'CMO' then bRetour := True                     // Choix des modules
  else if ((stAccueil = 'COC') and (JaiLeDroitTag( 52110))) then CPLanceFiche_CPCONSGENE()         // Consultation des comptes
  else if ((stAccueil = 'COE') and (JaiLeDroitTag(7602))) then OperationsSurComptes('','','')    // Consultation des écritures
  else bRetour := True;
  Result := bRetour;
end;


// GC - 05/03/2002
Procedure DispatchPCL10 ;
Var i : Integer ;
    St,Nom,Value : String ;
//    lCEtatsChaines : TCEtatsChaines;
BEGIN
  if ctxPCL in V_PGI.PGIContexte then
  begin
    for i:=1 to ParamCount do
    begin
      St:=ParamStr(i) ; Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
      Value:=UpperCase(Trim(St)) ;

      if Nom='/ETEBAC' then
       begin
        // Lancement d' une fiche pour mettre en inside, sinon le premier état chainé de s' imprime pas
        // Voir ligne 318 dans QRBalgen " if PP = nil then "
        {$IFNDEF EAGLCLIENT}
        CPLanceFiche_SaisieTresorerie( 'DP');
        {$ENDIF}
        // Fermeture de l'application
        Application.ProcessMessages;
        FMenuG.ForceClose := True ;
        FMenuG.Close ;
        Exit;
       end;

      if Nom='/ETATSCHAINES' then
      begin
        // Lancement d' une fiche pour mettre en inside, sinon le premier état chainé de s' imprime pas
        // Voir ligne 318 dans QRBalgen " if PP = nil then "
        {$IFNDEF EAGLCLIENT}
        CPLanceFiche_EtatsChaines(Value + ';CP');
        {$ENDIF}
        // Fermeture de l'application
        Application.ProcessMessages;
        FMenuG.ForceClose := True ;
        FMenuG.Close ;
        Exit;
      end;

      // Traitement par lot de l'alignement des standards
      if Nom = '/ALIGNESTD' then
      begin
        CPAlignementMultiDossier( Value );
        Application.ProcessMessages;
        FMenuG.ForceClose := True ;
        FMenuG.Close ;
      end;
    end;

    // Demande d'alignement sur le standard de référence
    if (GetParamSoc('SO_CPOKMAJPLAN')) and (GetParamSoc('SO_NUMPLANREF') > 0) then
    begin
      if StrToDateTime(GetParamSoc('SO_CPDATEDERNMAJPLAN')) <
         StrToDateTime(GetColonneSQl('STDCPTA', 'STC_DATEMODIF', 'STC_NUMPLAN = ' + IntToStr(GetParamSoc('SO_NUMPLANREF')))) then
      begin
        if PgiAsk('Le standard de référence du dossier à été modifié.' + #13 + #10 +
                  'Voulez vous effectuer un alignement ?', 'Alignement des standards') = MrYes then
        begin
          V_PGI.ZoomOLE := TRUE;
          CPLanceFiche_AlignementSurStandard( True );
          V_PGI.ZoomOLE := False;
        end;
      end;
    end;
  end;
END ;

Procedure LanceEtatLibreS5 ;
Var
  CodeD : String;
BEGIN
  CodeD:=Choisir('Choix d''un état libre','MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="UCO" AND MO_MENU="X"', '', False, False, 999999112, '');
  if CodeD<>'' then LanceEtat('E','UCO',CodeD,True,False,False,Nil,'','',False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Import de la fonction de S7 pour le S5 pack avancé
Mots clefs ... : PACK AVANCE;ETATS LIBRES
*****************************************************************}
Procedure LanceEtatLibreS7 ( NatureEtat : String ) ;
Var
  CodeD : String;
BEGIN
  CodeD:=Choisir('Choix d''un état libre','MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="'+NatureEtat+'" AND MO_MENU="X"','');
  if CodeD<>'' then LanceEtat('E',NatureEtat,CodeD,True,False,False,Nil,'','',False);
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

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
var
{$IFNDEF EAGLCLIENT}
  {$IFDEF NETEXPERT}
    NoSEqNet            : integer;
    Fileini,TypeExport  : string;
  {$ENDIF}
{$ENDIF}
    stAExclureFavoris   : string;
BEGIN
  // MODIF PACK AVANCE
  if not V_PGI.OutLook then // Pour empêcher la mise en place inside
    PRien := nil ;          //  en mode non outlook

  case Num of
    10 : begin
{$IFNDEF EAGLCLIENT}
      {$IFDEF SCANGED}
      if (VH^.OkModGed) then
      begin
        if V_PGI.RunFromLanceur then
          InitializeGedFiles(V_PGI.DefaultSectionDbName, MyAfterImport)
        else
          InitializeGedFiles(V_PGI.DbName, MyAfterImport);
      end;
      {$ENDIF}
      if not SupprimeEcartConversionEuro then begin
        FMenuG.ForceClose := TRUE;
        PostMessage(FmenuG.Handle,WM_CLOSE,0,0) ;
        RetourForce:=true;
        FermeHalley:=true;
        exit;
      end;

      ChargeHalleyUser ;

      ControleLettrageEnSaisie ;
(*
      if PGI_IMPORT_BOB('CCS5') = 1 then begin
        stMessage := 'LE SYSTEME A DETERMINE UNE MAJ DE MENU';
        stMessage := stMessage +#13#10 +'POUR ETRE PRISE EN COMPTE';
        stMessage := stMessage +#13#10 +'LE LOGICIEL VA SE FERMER AUTOMATIQUEMENT';
        stMessage := stMessage +#13#10 +'VOUS DEVEZ RELANCER L''APPLICATION';
        PGIInfo(stMessage+' : '+V_PGI.DBName,'MAJ MENU');
        FMenuG.ForceClose := TRUE;
        PostMessage(FmenuG.Handle,WM_CLOSE,0,0) ;
        RetourForce:=true;
        FermeHalley:=true;
        exit;
      end;
*)
      { Affichage de l'écran d'accueil }
      if not (ctxPCL In V_PGI.PGIContexte) then
        if AffichageEcranAccueil then RetourForce := True;
{$ENDIF}
      UpdateComboslookUp ; 
      // MODIF PACK AVANCE
      {$IFDEF TT}
        TripoteStatus ;
        DispatchPCL10 ;
      {$ENDIF}
      if EstComptaPackAvance then
        TripoteStatus(GetStRevis)
      else begin
        TripoteStatus ;
        DispatchPCL10 ;
      end ; // FIN MODIF PACK AVANCE
      // Route sur la messagerie PGI
{$IFDEF EAGLCLIENT}
{$ELSE}
      if FMDispS3.SeriaMessOK then OnAglMailForm := MyAglMailForm;
{$ENDIF}
    end;

    11 : ;// Après déconnexion

    12 : BEGIN
      {$IFNDEF EAGLCLIENT}
        // Interdiction de lancer en direct
        if (Not V_PGI.RunFromLanceur) and (V_PGI.ModePCL='1') then
        begin
          FMenuG.FermeSoc;
          FMenuG.Quitter;
          exit;
        end;
      {$ENDIF}

      InitSerialisation;
      if V_PGI.ModePCL='1' then
      begin
        V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
        RenseignelaSerie(ExeCCS3) ;
      end;
      if ctxPCL in V_PGI.PGIContexte then BEGIN
        {$IFNDEF EAGLCLIENT}
        FMenuG.bUser.Hint := TraduireMemoire('Exercice de référence');
        FMenuG.bUser.Onclick := FMDispS3.ExercClick;
        {$ENDIF}
        V_PGI.CodeProduit:='0011';
        V_PGI.QRPDF := True;
      END ;
    END ;

    13 :  begin
{$IFNDEF EAGLCLIENT}
        if ctxPCL in V_PGI.PGIContexte then begin
          InitialiserTableCompta;
          if VH^.OkModGed then FinalizeGedFiles;
          {$IFDEF NETEXPERT} // lien Netexpert
        // Reprise Net expert ajout me 15/09/2004
          if (IsDossierNetExpert(V_PGI.NoDossier, NoSEqNet)) and (ExisteSQL('SELECT E_IO FROM ECRITURE WHERE E_IO="X" or E_PAQUETREVISION=1')) then
             if PGIAsk('Vous avez effectué des modifications sur votre dossier, voulez-vous synchroniser avec '+VH^.CPLienGamme +' ?','Envoyer')=mrYes then
             begin
                Fileini := GetEnvVar('TEMP') + '\COMSX.ini';
                if NoSEqNet >= 1 then TypeExport := 'SYN'
                else TypeExport := 'DOS';
                GenereIniExport(Fileini, 'PG'+V_PGI.NoDossier+'_'+IntTostr(NoSEqNet)+'.TRA', TypeExport);
                ExportDonnees(Fileini+';EXPORT;Minimized',TRUE) ;
             end;
          {$ENDIF}
        end;
{$ENDIF}
    end;

    15 : ;

    16 : begin
{$IFNDEF EAGLCLIENT}
        if (not (ctxPCL in V_PGI.PGIContexte)) then
          ExportS7ImportS5Confidentialite;
        PGI_IMPORT_BOB('CCS5');
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
            else if ((ctxCompta in V_PGI.PGIContexte) and (V_PGI.RunFromLanceur)) then SetFlagAppli('CCS3.EXE',True);
    end;

//======================
//==== SPECIFIQUES  ====
//======================
{$IFNDEF EAGLCLIENT}
{$IFDEF CERIC}
     // Menu Spécifique CERIC
     9050 : AFFEnCours ;       // Affaire en cours
     9100 : GeneMvt ;          // Générations des mouvements
     9150 : ImprimEtatMarge ;  // Impression des affaires en cours
     9200 : ModifEcrCourante ; // Modifications des écritures courantes
     9250 : ReInitAff ;        // Réinitialisation des affaires
     9300 : DelMvt ;           // Destructions des mouvements
     9350 : ImprimBrouillard ; // Impression du brouillard
     9400 : CreationDossier ;  // Création des dossiers
{$ENDIF}

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
    7244 : MultiCritereMvt(taCreat,'N',False) ;       // Création
    7241 : SelectGuide('','','','',True) ;            // saisie guidée
    7256 : MultiCritereMvt(taConsult,'N',False) ;     // Consultation
    7259 : MultiCritereMvt(taModif,'N',False) ;	      // Modification
    7262 : DetruitEcritures('N',False) ;              // Suppression Ecritures Pieces
    7263 : DetruitEcritures('N',True) ;               // Suppression Ecritures Bordereaux
    7005 : SaisieFolio(taModif) ;
    52116 : CPLanceFiche_CPMULCUTOFF('') ;
    7260 : ModifSerieTableLibreEcr('E',[tpReel]) ; // MODIF PACK AVANCE {JP 10/06/05 : migration en eAGL}
{$IFDEF EAGLCLIENT}
    //RR En test
    7247 : PrepareSaisTres ;
    //SG6 06/01/05 en test
    7245 : PrepareSaisTresEffet ;
    7264 : CCLanceFiche_ModifEntPie;
{$ELSE}
    7245 : PrepareSaisTresEffet ;
    7247 : PrepareSaisTres ;
    7264 :  if ctxPCL in V_PGI.PGIContexte Then CPLanceFiche_VisuTresorerie  //LG* 06/05/2002
                                           Else CCLanceFiche_ModifEntPie;
    7268 : Brouillard('N') ;
{$ENDIF}

    // Modules SAISIE paramètrables
    9105 : MultiCritereMvt(taCreat,'M',False) ;       // Création
    9110 : MultiCritereMvt(taConsult,'M',False) ;     // Consultation
    9115 : MultiCritereMvt(taModif,'M',False) ;	      // Modification

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
{$IFDEF EAGLCLIENT}
    7281 : CPLanceFicheValideEnReel('S');
{$ELSE}
    7281 : ValideEnReel('S') ;
    7289 : Brouillard('S') ;
{$ENDIF}

    // Situation
    7298 : MultiCritereMvt(taCreat,'U',False) ;
    7295 : MultiCritereMvt(taConsult,'U',False) ;
    7301 : MultiCritereMvt(taModif,'U',False) ;
    7304 : DetruitEcritures('U',False) ;
    7305 : ModifSerieTableLibreEcr('E',[tpSitu]) ; // MODIF PACK AVANCE {JP 14/06/05 : Migration eAgl}

    // Prévisions
{$IFDEF EAGLCLIENT}
    7313 : MultiCritereMvt(taCreat,'P',False) ;     // MODIF PACK AVANCE
    7310 : MultiCritereMvt(taConsult,'P',False) ;   // MODIF PACK AVANCE
    7316 : MultiCritereMvt(taModif,'P',False) ;     // MODIF PACK AVANCE
    7319 : DetruitEcritures('P',False) ;            // MODIF PACK AVANCE
{$ELSE}
    7313 : MultiCritereMvt(taCreat,'P',False) ;     // MODIF PACK AVANCE
    7310 : MultiCritereMvt(taConsult,'P',False) ;   // MODIF PACK AVANCE
    7316 : MultiCritereMvt(taModif,'P',False) ;     // MODIF PACK AVANCE
    7319 : DetruitEcritures('P',False) ;            // MODIF PACK AVANCE
{$ENDIF}
    7317 : ModifSerieTableLibreEcr('E',[tpPrev]) ;  // MODIF PACK AVANCE {JP 14/06/05 : Migration eAgl}

    // Révision
{$IFDEF EAGLCLIENT}
    7649 : if OkRevis then MultiCritereMvt(taCreat,'R',False) ;     // MODIF PACK AVANCE
    7658 : if OkRevis then MultiCritereMvt(taConsult,'R',False) ;   // MODIF PACK AVANCE
    7661 : if OkRevis then MultiCritereMvt(taModif,'R',False) ;     // MODIF PACK AVANCE
    7662 : if OkRevis then CPLanceFicheValideEnReel('R');           // MODIF PACK AVANCE
    7664 : if OkRevis then DetruitEcritures('R',False) ;            // MODIF PACK AVANCE
    //7670 : if OkRevis then Brouillard('R') ;                        // MODIF PACK AVANCE
    //7682 : if OkRevis then LanceRevision ;      // MODIF PACK AVANCE
{$ELSE}
    7649 : if OkRevis then MultiCritereMvt(taCreat,'R',False) ;     // MODIF PACK AVANCE
    7658 : if OkRevis then MultiCritereMvt(taConsult,'R',False) ;   // MODIF PACK AVANCE
    7661 : if OkRevis then MultiCritereMvt(taModif,'R',False) ;     // MODIF PACK AVANCE
    7662 : if OkRevis then ValideEnReel('R') ;                      // MODIF PACK AVANCE
    7664 : if OkRevis then DetruitEcritures('R',False) ;            // MODIF PACK AVANCE
    7670 : if OkRevis then Brouillard('R') ;                        // MODIF PACK AVANCE
    7682 : if OkRevis then LanceRevision ;      // MODIF PACK AVANCE
{$ENDIF}

    // Abonnements
    7328 : GenereAbonnements(True) ;
    7331 : GenereAbonnements(False) ;
    7337 : ParamAbonnement(TRUE,'',taModif) ;
    7340 : ReconductionAbonnements ;
    7346 : ListeAbonnements ;

    // Analytiques
    7113 : CCLanceFiche_MULGeneraux('SERIEANA;7115000;1');
    7367 : MultiCritereAna(taCreat) ;
    7364 : MultiCritereAna(taConsult) ;
    7370 : MultiCritereAna(taModif) ;
    7382 : CPLanceFiche_ReventilAna;
    7379 : VidageInterSections ;             // MODIF PACK AVANCE
    7371 : ModifSerieTableLibreEcr('A',[]) ; // MODIF PACK AVANCE {JP 14/06/05 : Migration eAgl}
    // fa 2005.08.31
    // 7373 : CPLanceFicheMulSupprAna;  //    7373 : DetruitAnalytiques ;
{$IFNDEF EAGLCLIENT}
    7385 : BrouillardAna('N') ;
    {JP 16/05/05 : FQ 15865}
    7381 : if VH^.AnaCroisaxe then CC_LanceFicheReimpAna
                              else ReImputAnalytiques;
{$ENDIF}

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
{$IFDEF EAGLCLIENT}
{$ELSE}
    7598 : LanceCoutTiers('',0) ;
{$ENDIF}

    // Editions de suivi
    7547 : CPLanceFiche_BalanceAgee;
(*
    7556: if ctxPCL in V_PGI.PGIContexte then AGLLanceFiche('CP', 'EPBGLGEN', '', 'BVE', 'BVE') // Balance ventilée
      else CPLanceFiche_BalanceVentilee;  // FQ 14149 - 21/01/2005
      *)
      7556 : CPLanceFiche_BalanceVentilee;  // FQ 14149 - 21/01/2005
    7560 : CPLanceFiche_CPGLAUXI(''); // GLAuxiliaireL ; // FQ 13981 / SBO 29/09/2004
    7541 : AGLLanceFiche('CP', 'EPECHEANCIER','','ECH','ECH');
{$IFDEF EAGLCLIENT}
{$ELSE}
    7535 : JustSolde ;
    7550 : GLivreAge ;
    7559 : GLVentile ;
{$ENDIF}

    // Tiers payeurs
    7231 : AGLLanceFiche('CP', 'EPTIERSP', '', 'TFP', 'TPF') ; // Tiers facturé par payeur
    7232 : AGLLanceFiche('CP', 'EPTIERSP', '', 'TFA', 'TPF') ; // Tiers facturé avec payeur
{$IFDEF EAGLCLIENT}

{$ELSE}
    7233 : BrouillardTP ;
    7234 : BalanceAge2TP ;
    7236 : GLivreAgeTP ;
    7235 : BalVentileTP ;
    7237 : GLVentileTP ;
{$ENDIF}

    //Emissions
    7589 : ReleveCompte ;   // Relevés de comptes
    7592 : ReleveFacture ;  // Relevés de factures
{$IFDEF EAGLCLIENT}
    7585 : ModifEcheMP;
{$ELSE}
    7599 : ExportCFONBBatch(False,True,tslAucun) ;
    7586 : ExportCFONBBatch(True,True,tslAucun) ;
    7595 : ExportCFONBBatch(True,False,tslAucun) ;
    7594 : AGLLanceFiche('CP', 'RLVEMISSION', '', '', '') ;
    7596 : ExportCFONBBatch(False,False,tslTraite) ;
    7597 : ExportCFONBBatch(False,False,tslBOR) ;
    7585 : ModifEcheMP ;
{$ENDIF}

    // Autres éditions
{$IFDEF EAGLCLIENT}

{$ELSE}
   52701 : BalanceAge;
   52702 : BalVentile;
   52705 : GLAuxiliaireL ;
{$ENDIF}

//==== Editions (13) ====
    // Journaux
    {JP 23/06/05 : Uniformisation des journaux PGE - PCL
                   Le menu est renommé en journal des écritures et non plus journal divisionnaire}
    7394 : CPLanceFiche_CPJALECR ; // 26/11/2002 nv journal des ecritures
{$IFDEF EAGLCLIENT}
{$ELSE}
    {JP 23/06/05 : Uniformisation des journaux PGE - PCL
    7397 : CumulPeriodique(fbJal,neJalCentr) ; cf Menu 7411 (voire le menu 7412, si l'on veut par période
    7403 : CumulPeriodique(fbJal,neJalPer) ; cf Menu 7413
    7406 : CumulPeriodique(fbJal,neJaG) ; cf Menu 7423}
    7409 : JalCpteGe  ;
{$ENDIF}

    // Grand-livres
    7415 : CPLanceFiche_CPGLGENE('');
    7418 : CPLanceFiche_CPGLAUXI('');
    7419 : CPLanceFiche_EtatGrandLivreSurTables; 	// Grand Livre sur table libre
    7427 : CPLanceFiche_CPGLGENEPARAUXI;
    7430 : CPLanceFiche_CPGLAUXIPARGENE;
    7421 : CPLanceFiche_CPGLANA ;
    7436: CPLanceFiche_CPGLGENEPARANA;
    7439: CPLanceFiche_CPGLANAPARGENE;


    // Balances
    7445 : CPLanceFiche_BalanceGeneral; 												// Balance générale eAGL
    7448 : CPLanceFiche_BalanceAuxiliaire; 			 								// Balance auxiliaire eAGL
    7429 : CPLanceFiche_EtatBalanceSurTables; 									// Balance sur table libre
    7451 : CPLanceFiche_BalanceAnalytique ;                     // Balance analytique
    7457 : CPLanceFiche_BalanceGenAuxi ;
    7460 : CPLanceFiche_BalanceAuxiGen ;
    7466 : CPLanceFiche_BalanceGenAnal ;
    7469 : CPLanceFiche_BalanceAnalGen ;

    {$IFDEF EAGLCLIENT}
    {JP 24/06/05 : Migration eAgl des Cumuls périodiques}
    7478 : CumulPeriodique(fbGene);
    7481 : CumulPeriodique(fbAux );
    7484 : CumulPeriodique(fbSect);
    {$ELSE}
    7478 : CumulPeriodique(fbGene,neJalRien) ;
    7481 : CumulPeriodique(fbAux,neJalRien) ;
    7484 : CumulPeriodique(fbSect,neJalRien) ;
    {$ENDIF}

    // Etats de synthèses
    7440 : EtatPCL (esCR)  ; // Document de synthèse - Résultat
    7441 : EtatPCL (esBIL) ; // Document de synthèse - Actif
    7442 : EtatPCL (esSIG) ; // Document de synthèse - SIG
    7446 : EtatPCL (esCRA) ; // Document de synthèse analytique - Résultat
    7449 : EtatPCL (esBILA) ; // Document de synthèse analytique - Bilan
    7447 : EtatPCL (esSIGA) ; // Document de synthèse analytique - SIG
    7444 : CPLanceFiche_EtatsChaines(';CP');
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
    7456 : LanceEtatLibreS5 ;

    // Autres éditions
{$IFDEF EAGLCLIENT}

{$ELSE}
   52585 : JalDivisio;
  {JP 23/06/05 : Uniformisation des journaux PGE - PCL : déplacement des anciens modules dans la
                 rubrique anciens états}
   52600 : CumulPeriodique(fbJal,neJalCentr);
   52601 : CumulPeriodique(fbJal,neJalPer)  ;
   52602 : CumulPeriodique(fbJal,neJaG)     ;
   {JP 26/07/05 : FQ 15231 : branchement des anciens GL Gene / Anal et Anal / Gene}
   52611 : GLGESE;
   52612 : GLSEGE;
   {JP 26/07/05 : Déplacement de ces lignes qui étaient situées 100 lignes plus haut}
   52586 : JalDivisio; // Journal des écritures (PCL)
   52592 : GLAuxiGen ; // Grand-livre auxiliaire par général (PCL)
   52593 : GLGenAuxi ; // Grand-livre général par auxiliaire (PCL)

   52583 : GLGeneral;
   52584 : GLAuxiliaire;
   52581 : BalanceGeneComp;
   52582 : BalanceAuxi;
   52587 : EtatPointage ;
   52588 : EtatRappro ;
   52589 : JustifPointage ;
   52594 : BLAnal ;
   52595 : BalGeAu ;
   52596 : BalAuGe ;
   52597 : BLGESE ;
   52598 : BLSEGE ;

{$ENDIF}

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
    8621 : CPLanceFiche_CPGLGENEM('');
    8622 : CPLanceFiche_CPGLAUXIM('');
    8623 : ;// GLAnal ; => Attente du passage en etat AGL
    8624 : CPLanceFiche_EtatGrandLivreSurTablesM; 	  // Grand Livre sur table libre
    8625 : CPLanceFiche_CPGLGENEPARAUXIM;
    8626 : CPLanceFiche_CPGLAUXIPARGENEM;
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
    52110 : CPLanceFiche_CPCONSGENE();      // Consultation des comptes généraux
    52120 : CPLanceFiche_CPCONSAUXI();      // Consultation des comptes auxiliaires

    7261 : Centralisateur(taModif, 'N', FALSE) ;
    7676 : CC_LanceFicheReimp;                        // Ré-imputations
{$IFDEF EAGLCLIENT}
    7677 : CPLanceFiche_Extourne;
{$ELSE}
    7677 : If Not EstSpecif('51189') Then CPLanceFiche_Extourne
                                     else LanceExtourne ;
    7719 : TLAuxversTLEcr ;
{$ENDIF}

    // Pointage
    7604 : CPLanceFiche_PointageMul ( );
    7616 : CC_LanceFicheEtatPointage ;        // Etat de pointage
// BPY : remplacement par le nouveau :    7619 : CC_LanceFicheEtatRappro ;          // Etat de rapprochement
    7619 : CC_LanceFicheEtatRapproDet ;          // Etat de rapprochement
    7622 : CC_LanceFicheJustifPointage ;      // Justificatif des soldes bancaires
{$IFDEF EAGLCLIENT}
{$ELSE}
// Plus de dépointage : passer par le pointage
//    7607 : If Not EstSpecif('51191') Then PointageCompteNew(FALSE) Else PointageCompte(false) ;
    7631 : ControlCaisse(fbGene,tzGCaisse,'') ;
{$ENDIF}

    // Validations
    7691 : ValidationPeriode(TRUE,FALSE);     // Validation par période
    7694 : ValidationPeriode(FALSE,FALSE);    // Dévalidation par période
    7700 : ValidationJournal(TRUE) ;          // Validation par journal
    7703 : ValidationJournal(FALSE) ;         // Dévalidation par journal
{$IFDEF EAGLCLIENT}
{$ELSE}
    7709 : LanceEdtLegal ;
{$ENDIF}

    // A-nouveaux
    7718 : OuvertureExo ; // Ouverture d'exercice
    7724 : MultiCritereMvt(taCreat,'N',True) ;
    7727 : MultiCritereMvt(taConsult,'N',True) ;
    // Clôtures
    7736 : CloPer;                          // Clôture périodique
    7737 : AnnuleCloPer;                    // Annulation de clôture périodique
    7742 : begin ClotureComptable(FALSE) ; {RetourForce := True;} end;      // Clôture provisoire //SG6 06/12/2004 FQ 14987
    7745 : begin AnnuleclotureComptable(FALSE); {RetourForce := True;} end; // Annulation de clôture provisoire //SG6 06/12/2004 FQ 14987
    7751 : If ctxPCL in V_PGI.PGIContexte then CPLanceClotureDefinitive
           else If LanceAuditPourCloture Then FermeHalley:=TRUE ;
{$IFDEF EAGLCLIENT}
{$ELSE}           
    7757 : JalDivisioAno ;
    7760 : BalanceAno ;
    7766 : JalDivisioClo ;
    7769 : BalanceClo ;
    77120 : VisuEditionLegale(TRUE) ; // MODIF PACK AVANCE
{$ENDIF}
    7683 : ResultatDeLexo ;

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
    7772 : CPLanceFiche_RecepETEBAC3; //
    7773 : CPLanceFiche_RecupReleve;  // Intégration des relevés
    7774 : CP_LancePointageAuto;
    7776 : CPLanceFiche_DPointageMul;
    7777 : CC_LanceFicheEtatRapproDet;

    // Révision
    54710 : ParamTable('CPPLANREVISION', taCreat, 0, PRien, 3, 'Plans de révisions') ;
    54720 : MultiCritereCycleRevision;
    54731 : ControleRubrique('CYCLE');  // Contrôle du paramétrage
    54732 : ControleRubrique('LIASSE'); // Contrôle du paramétrage

    // ICC
    96110 : CPLanceFiche_ICCGENERAUX;      // Comptes courants
    96130 : CPLanceFiche_FicheICCGENERAUX( GetParamSoc('SO_ICCCOMPTECAPITAL') ); // Compte de capital
    96140 : CPLanceFiche_ICCELEMNATIONAUX; // Eléments nationaux

    // CRE (PCL)
{$IFDEF EAGLCLIENT}
{$ELSE}
    96210 : AGLLanceFiche('FP', 'FMULEMPRUNT', '', '','');
    96220 : LancerEditions;
{$ENDIF}

    // EDI
    {$IFDEF EDICOMPTA}
    96310 :  EchgImp_LanceAssistant ('BAL');
    96320 :  EchgImp_LanceAssistant ('GLJ');
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
    7815 : AGLLanceFiche('CP', 'CPRUBTLIB', '', '', '') ;{... des tables libres}
{$IFDEF EAGLCLIENT}
    7811 : ParamTable('TTCONSTANTE', taCreat,0, PRien, 3, 'Constantes'); // Constantes

    //SG6 FQ 15288 21/01/05 7815 : AGLLanceFiche('CP', 'CPRUBTLIB', '', '', '') ;{... des tables libres}
{$ELSE}
    // Maquettes d'édition / Etats
    69110 : AGLLanceFiche('CP','STDMAQ','CR','','ACTION=MODIFICATION;CR');   // Compte de résultat (PCL)
    69310 : AGLLanceFiche('CP','STDMAQ','BIL','','ACTION=MODIFICATION;BIL'); // Bilan (PCL)
    69210 : AGLLanceFiche('CP','STDMAQ','SIG','','ACTION=MODIFICATION;SIG'); // SIG (PCL)

    7811 : Constantes ;
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
{$IFDEF EAGLCLIENT}

{$IFDEF PORTAGEBUDGET}
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
    15321 : BalBudteGen;               // Balances / Par comptes
    15322 : BalBudteSec;               // Balances / Par sections
    15323 : BalBudteGenSec;            // Balances / Par comptes et par sections
    15324 : BalBudteSecGen;            // Balances / Par sections et par comptes
    15331 : BalBudecGen;               // Ecart budgétaires / Par comptes
    15332 : BalBudecSec;               // Ecart budgétaires / Par sections
    15333 : BalBudecGenSec;            // Ecart budgétaires / Par comptes et par sections
    15334 : BalBudecSecGen;            // Ecart budgétaires / Par sections et par comptes
    15282 : BrouillardBud('N');        // Brouillard / Brouillard par comptes budgétaires
//    15284 : BrouillardBudAna('N');     // Brouillard / Brouillard par sections budgétaires
//    15340 : EcartParCategorie;         // Etat multi-budgets
    15350 : LanceEtatLibreS7('UBU');   // Etats libres

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
    15425 : ImpRubrique('',False,True);           // Impression des rubriques

    // Paramètres
    1496 : YYLanceFiche_Souche('BUD','','ACTION=MODIFICATION;BUD'); // Compteurs
    1498 : ParamStructureBudget(PRien);                             // Catégories budgétaires
    1530 : ParamRepartBUDGET('',False);                             // Clés de répartition mensuelles
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
{$ENDIF}

{$ELSE}
    // Ecritures
    15211 : MultiCritereMvtBud(taCreat,'G') ;
    15213 : MultiCritereMvtBud(taCreat,'S') ;
    15217 : MultiCritereMvtBud(taCreat,'GS') ;
    15219 : MultiCritereMvtBud(taCreat,'SG') ;
    15221 : MultiCritereMvtBud(taConsult,'') ;
    15230 : MultiCritereMvtBud(taModif,'') ;
    15225 : VisuConsoBudget('G','') ;
    15227 : VisuConsoBudget('S','') ;
    15240 : DetruitBudgets ;

    // Structures
    15115 : MulticritereBudgene(taModif); // Comptes
    15135 : MulticritereBudsect(taModif); // Sections
    15155 : MulticritereBudjal(taModif);  // Budgets
    15119 : SuppressionCpteBudG ;
    15139 : SuppressionCpteBudS ;
    15159 : SuppressionJournauxBud ;
    15125 : PlanBudget('',FALSE) ;
    15145 : PlanBudSec('','',FALSE) ;
    15165 : PlanBudJal('','',False) ;
    15117 : MulticritereBudgene(taModifEnSerie) ;
    15137 : MulticritereBudsect(taModifEnSerie) ;
    15157 : MulticritereBudjal(taModifEnSerie) ;
    15118 : ModifSerieTableLibre(fbBudGen) ;     // MODIF PACK AVANCE
    15138 : ModifSerieTableLibre(fbBudSec1) ;    // MODIF PACK AVANCE

    // Editions
    15321 : BalBudteGen ;
    15322 : BalBudteSec ;
    15323 : BalBudteGenSec ;
    15324 : BalBudteSecGen ;
    15331 : BalBudecGen ;
    15332 : BalBudecSec ;
    15333 : BalBudecGenSec ;
    15334 : BalBudecSecGen ;
    15282 : BrouillardBud('N') ;
    15284 : BrouillardBudAna('N') ;
    15340 : EcartParCategorie ;        // MODIF PACK AVANCE
    15350 : LanceEtatLibreS7('UBU') ;  // MODIF PACK AVANCE

    // Traitements
    15250 : MultiCritereValBud(True) ;
    15260 : MultiCritereValBud(False) ;
    15270 : RecopieBudgetSimple ;
    15271 : RecopieBudgetMultiple ;
    15272 : RevisionBudgetaireRealiser ;
    15161 : OuvreFermeCpte(fbBudJal,False) ; // MODIF PACK AVANCE
    15163 : OuvreFermeCpte(fbBudJal,True) ;  // MODIF PACK AVANCE

    // Reporting
    15410 : MultiCritereRubriqueV2(taModif,True) ;
    15430 : ControleRubriqueBud ;
    15440 : TotalRubriquesV2(True) ;
    15421 : SuppressionRubriqueV2(True) ;
    15425 : ImpRubrique('',False,True) ;

    // Paramètres
    1496 : YYLanceFiche_Souche('BUD','','ACTION=MODIFICATION;BUD'); //Compteurs
    1498 : ParamStructureBudget ;
    1530 : ParamRepartBUDGET('',False) ;
    15412 : ParamCroiseCompte ;
    1501 : CodeGeneVersCodeBudg ;
    1502 : CreationCodBudSurRupture(fbGene) ;
    1504 : CreationDesCodesBudget(fbBudGen) ;  // MODIF PACK AVANCE
    1522 : CreationCodBudSurRupture(fbAxe1) ;
    1528 : GenerationSectionBudgetaire ;   {FP FQ15648}
    1524 : CreationDesCodesBudget(fbBudSec1) ; // MODIF PACK AVANCE
    1529 : GenSecBudCategorie ;
    15411 : CreationRubrique ;
    15126 : ControleBudget(fbBudGen) ;
    15146 : ControleBudget(fbAxe1) ;
{$ENDIF}

//===========================================
//==== Amortissement (20-PGE) et (2-PCL) ====
//===========================================
{$IFDEF AMORTISSEMENT}
    // Structures
    2111 : AMLanceFiche_ListeDesImmobilisations ( '' );
    2115 : AMLanceFiche_ListeDesImmobilisations ( '' , True);
    2150 : CPLanceFiche_EtatsChaines(';AM');

    // Editions
    2549 : AMLanceFiche_Edition ( 'IPM', 'IPM' );
    2542 : AMLanceFiche_Edition ( 'IAC', 'IAC' );
    2541 : AMLanceFiche_Edition ( 'ILS', 'ILS' ); // PCL
//    2546 : AMLanceFiche_Edition ( 'AMN', 'AMN' );
    2543 : AMLanceFiche_Edition ( 'ISO', 'ISO' );
    2544 : AMLanceFiche_Edition ( 'ITE', 'ITE' );
    20222 : AMLanceFiche_Edition ( 'IMU', 'IMU' );
    20232 : AMLanceFiche_Edition ( 'ITN', 'ITN' );
    20234 : AMLanceFiche_Edition ( 'ITF', 'ITF' );
    20236 : AMLanceFiche_Edition ( 'ITR', 'ITR' );
    20238 : AMLanceFiche_Edition ( 'ITO', 'ITO' );
    20239 : AMLanceFiche_Edition ( 'IRE', 'IRE' );        
    20252 : AMLanceFiche_Edition ( 'ITP', 'ITP' );
    2547 : AMLanceFiche_Edition ( 'IPR', 'PRE' );
    2548 : AMLanceFiche_Edition ( 'IPR', 'PRF' );
    20272 : AMLanceFiche_Edition ( 'IVE', 'IVE' );
    20274 : AMLanceFiche_Edition ('IVL','IVL' );
    2180 : AMLanceFiche_StatPrevisionnel;

{$IFNDEF EAGLCLIENT}

    // Anciennes éditions
    2505 : LanceEditionImmo ('FIM');
    2521 : LanceEditionImmo ('DOT');
    2522 : LanceEditionImmo ('DEC');
    2523 : LanceEditionImmo ('DFI');
    2524 : LanceEditionImmo ('DDE');
    2512 : LanceEditionImmo ('INV');
    2511 : LanceEditionImmo ('ACS');
    2513 : LanceEditionImmo ('SIM');
    2514 : LanceEditionImmo ('STS');
    2516 : LanceEditionImmo ('PMV');
    2515 : LanceEditionImmo ('MUT');
    2535 : LanceEditionImmo ('DCB');
    2536 : LanceEditionImmo ('ECB');
    2539 : LanceEditionImmo ('PTP');
    2531 : LanceEditionImmo ('PRE');
    2532 : LanceEditionImmo ('PRF');
    2540 : LanceEditionImmo ('ETI');
{$ENDIF}
    // Traitements
    2410 : AMLanceFiche_ListeDesIntegrations ('');
    2415 : EtatRapprochementCompta ;
    2416 : AfficheClotureImmo ;
    2417 :
      begin
        AMLanceFiche_AnnulationCloture;
        RetourForce:=true;
      end;
    2418 : UpdateBaseImmo ;    //ajout TD le 8/6/2001 menu Recalcul des plans

    // Paramètres
    2340 : ParamSociete(False,'','SCO_IMMOBILISATIONS','',RechargeParamSoc,ChargePageImmo,SauvePageImmo,Nil,2340000);
    2210 : AMLanceFiche_ComptesAssocies ( taModif );
    2342 : ParamTable('TILIEUGEO',taCreat,2342000,PRien) ;
    2345 : ParamTable('TIMOTIFCESSION',taCreat,2345000,PRien) ;
    2347 : SaisieDesCoefficientsDegressifs ;
    2348 : ParamTable('AMREGROUPEMENT',taCreat,0,PRien);
    // Utilitaires
//    2370 : ControleFichiersImmo ;
//    2375 : RepareFichiersImmo ;
    2370 : LanceVerificationAmortissement;
    2375 : LanceVerificationAmortissement;

    {$IFNDEF EAGLCLIENT}
    20511 : SavImo2PGI ('',False );
    20512 : ExporteAmortissement ;
    {$ENDIF}
{$ENDIF}

//======================
//==== Analyses (8) ====
//======================
    // Cubes
    8305  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','') ;
    8307  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','CONTREPARTIE') ;
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

{$IFDEF EAGLCLIENT}
    7133 : CPLanceFiche_PLANGENE('');     // Impression / Comptes généraux
    7166 : CPLanceFiche_PLANAUXI('');     // Impression / Comptes auxiliaires
    7199 : CPLanceFiche_PLANSECT('', ''); // Impression / Sections analytiques
    7229 : CPLanceFiche_PLANJOUR('');     // Impression / Journaux
{$ELSE}
    7133 : PlanGeneral('', false) ;
    7166 : PlanAuxi('',False) ;
    7199 : PlanSection('','',FALSE) ;
    7229 : PlanJournal('',FALSE) ;
{$ENDIF}

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
{$IFDEF EAGLCLIENT}
    1104: ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
      RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
{$ELSE}
    1104: ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
      RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
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

    // Compteurs
    1300 : YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT'); //Compteurs
    1301 : YYLanceFiche_Souche('REL','','ACTION=MODIFICATION;REL'); //Compteurs

    //Multidossier
    1750 : ParamRegroupementMultiDossier ;
    {$IFDEF MULTISOC}
    1755 : YYLanceFiche_Bundle( 'YY', 'YYBUNDLE', '', '', '') ;
    {$ENDIF MULTISOC}
    1760 : CPLanceFiche_LiensEtab('','','') ;


{$IFDEF EAGLCLIENT}
{$ELSE}
    1325 : CCLanceFiche_Correspondance('GE');	// Plan de correspondance Généraux
    1330 : CCLanceFiche_Correspondance('AU');	// Plan de correspondance Auxiliaire
    1345 : CCLanceFiche_Correspondance('A1');	// Plan de correspondance Axe analytique 1
    1350 : CCLanceFiche_Correspondance('A2');	// Plan de correspondance Axe analytique 2
    1355 : CCLanceFiche_Correspondance('A3');	// Plan de correspondance Axe analytique 3
    1360 : CCLanceFiche_Correspondance('A4');	// Plan de correspondance Axe analytique 4
    1365 : CCLanceFiche_Correspondance('A5');	// Plan de correspondance Axe analytique 5
    1426 : CCLanceFiche_Correspondance('ZG') ;
{$ENDIF}

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
    1430 : ParamGuide('','NOR',taModif) ;
    1435 : ParamScenario('','') ;

    1485 : ParamEtapeReg(True,'',True) ;
    1490 : ParamEtapeReg(False,'',True) ;
    1491 : CircuitDec ;
{$ENDIF}
    1425 : ParamLibelle ;          // Libellés automatiques
    1450 : ParamPlanAnal('');      // Analytiques / Structures analytiques
    1455 : ParamRepartCle;         {JP 08/06/05 : Répartition des clefs analytiques}
    1460 : ParamVentilType(PRien); // Analytiques / Ventilations types

    7574 : CCLanceFiche_ParamRelance('RTR');
    7577 : CCLanceFiche_ParamRelance('RRG');

    // Ecritures
{$IFDEF EAGLCLIENT}
    7352 : CPLanceFiche_MulGuide('','','');   // Guide
    7355 : CPLanceFiche_Scenario('','');      // Scénario
{$ELSE}
    7352 : ParamGuide('','NOR',taModif) ;
    7355 : ParamScenario('','') ;
{$ENDIF}

    // Banques
{$IFDEF TRSYNCHRO}
    {JP 10/03/04 : Ajout d'un ParamSoc qui détermine si l'on utilise la Tréso ...}
      1705 : if EstComptaTreso then TRLanceFiche_Banques('TR','TRCTBANQUE', '', '', '')
                               else FicheBanque_AGL('',taModif,0);
      1706 : TRLanceFiche_Agence('TR','TRAGENCE', '', '', '');
      1707 : TRLanceFiche_BanqueCP('TR','TRMULCOMPTEBQ','','','');
{$ELSE}
    1705 : FicheBanque_AGL('',taModif,0);                // Etablissements bancaires
{$ENDIF}
    1720 : ParamTable('ttCFONB',taCreat,500005,PRien);        // Codes CFONB
    1725 : ParamCodeAFB;                                 // Code AFB
{$IFDEF EAGLCLIENT}

{$ELSE}
    1730 : ParamTable('ttCodeResident',taCreat,500007,PRien) ;
    1732 : ParamTeletrans ;
{$ENDIF}

    // Documents
{$IFDEF EAGLCLIENT}

{$ELSE}
    1235 : EditDocument('L','REL','',TRUE) ;
    1240 : EditDocument('L','RLV','RLV',FALSE) ;
    1245 : EditDocument('L','RLC','',True) ;
    1247 : EditEtat('E','SAI','',True,nil,'','') ;
    1248 : EditEtat('E','BAP','',True,nil,'','') ;

    1267 : EditDocument('L','LCH','',True) ;
    1268 : EditDocument('L','LTR','',True) ;
//    1269 / // SUPPRIME
//    1266 / // SUPPRIME
    1257 : EditEtat('E','BOR','',True,nil,'','') ;

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
{$IFDEF EAGLCLIENT}
    3140 : if ctxPCL in V_PGI.PGIContexte then LanceAssistantComptaPCL(False);
{$ELSE}
    3130 : GestionSociete(PRien,@InitSociete,nil) ;
    3140 : if ctxPCL in V_PGI.PGIContexte then LanceAssistantComptaPCL(False) else LanceAssistSO ;
    3145 : RecopieSoc(PRien,True) ;
    3247 : CPLanceFiche_EnregStd;
    // PCL

    3222 : AGLLanceFiche('CP','REPRISESTAND','','','');      // INUTILISE
//    3246 : If BasculeCompta Then FermeHalley:=TRUE ;         // INUTILISE
    3249 : AGLLanceFiche('CP', 'RECUPDONNEES', '', '', '') ; // INUTILISE
{$ENDIF}
    3172 : YYLanceFiche_ProfilUser;
    3248 : CPLanceFiche_AlignementSurStandard;
    // Utilisateurs et accès
    3165  : FicheUserGrp ;
    3170 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;
    60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','6;8;9;11;12;13;14;16;17;18;20;26;27');
{$IFDEF EAGLCLIENT}
    3180 : ReseauUtilisateurs(False) ;
    3195 : ReseauUtilisateurs(True) ;
{$ELSE}
    3180 : ReseauUtilisateurs(False) ;
    3185 : VisuLog ;
    3195 : ReseauUtilisateurs(True) ;
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
        else begin
          stAExclureFavoris := '';
          TripotageMenuModePGE ( [9,11,13,16,14], stAExclureFavoris );
          TripotageMenuAmortissement (stAExclureFavoris );
          if ctxPCL in V_PGI.PGIContexte then
            ParamFavoris('12;2;8;17;18', stAExclureFavoris, False, True)
          else ParamFavoris('12;20;8;17;18', stAExclureFavoris, False, True);
        end ;
      end;
    3205 : CCLanceFiche_LongueurCompte ;
    7119 : CRazModePointage ;          // Changement du mode de pointage

    {$IFDEF TRSYNCHRO}
    3217 : Commun.CreerAgence; {JP 10/03/04 : mise ne place de l'initialisation des agences}
    {$ENDIF}
{$IFDEF EAGLCLIENT}
    {$IFDEF RR}
    3212 : CPLanceFiche_CPANOECR ;
    {$ENDIF}
{$ELSE}
    3210 : ControlFic ;
    {$IFDEF RR}
    3212 : CPLanceFiche_CPANOECR ;
    {$ENDIF}
    3215 : ReparationFic ;
    3220 : ReindexSoc ;
    3221 : CPLanceFiche_RubImport;
    56180 : CCLanceFiche_ModifEntPie; // Modification d'entête de pièces (PCL)
    56190 : CPLanceFiche_CPPurgeExercice;
{$ENDIF}

    // Comptable
    3227 :CPLanceGestionAna;
    3245 : AnnuleclotureComptable(TRUE);

//=======================
//=== Menu Pop Compta ===
//=======================
    25710 : CCLanceFiche_MULGeneraux('P;7112000');  // Comptes généraux
    25720 : CPLanceFiche_MULTiers('P;7145000');     // Comptes auxiliaires
    25730 : CPLanceFiche_MULSection('P;7178000');   // Sections analytiques
    25740 : CPLanceFiche_MULJournal('P;7205000');   // Journaux
    25750 : MultiCritereMvt(taConsult,'N',False);   // Consultation des écritures
    25755 : OperationsSurComptes('','','') ;  	    // Consultation des comptes
    25760 : ResultatDeLexo ;
    25770: ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
      RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);

//=======================
//=== Menu Pop Budget ===
//=======================
{$IFDEF EAGLCLIENT}

{$IFDEF PORTAGEBUDGET}
    26160 : MulticritereBudgene(taConsult);
    26165 : MulticritereBudsect(taConsult) ;
    26170 : MulticritereBudjal(taConsult) ;
    26175 : MultiCritereMvtBud(taConsult,'');
{$ENDIF PORTAGEBUDGET}
{$ELSE}
    26160 : MulticritereBudgene(taConsult) ;
    26165 : MulticritereBudsect(taConsult) ;
    26170 : MulticritereBudjal(taConsult) ;
    26175 : MultiCritereMvtBud(taConsult,'') ;
{$ENDIF}

    26180 : CCLanceFiche_MULGeneraux('C;7112000');  // Comptes généraux
    26185 : CPLanceFiche_MULSection('C;7178000');   // Sections analytiques
    26190: ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
      RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);


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

    // Déclaration de TVA (ajout SB) group -7660
{    7660 :	AGlLanceFiche ('CP','FICHEPARAMCA3','0','0','');  							// Paramètres généraux
	  7661 :  AGLLanceFiche('CP','MULINFOCA3','','','');										  // Liste des CA3
    7662 :	AGLLanceFiche('FIS','L0000_CONTROLE','','','CtrlTVA2001;TRUE'); // Contrôle de conformité du CA3 en cours
    7663 :	CA3PrintDeclaTVA (FALSE);	  																		// Etats CA3
    7664 :	CA3PrintDeclaTVA (TRUE);																				// Etats vierge CA3
    7665 :	CPLanceFiche_TVAEDI;																						// Dématérialisation}

    16110 : FicheBanque_AGL('',taModif,0);  // INUTILISE
    16120 : FicheBanqueCP('',taModif,0);    // INUTILISE

    45101 : ParamTable('CPNATURETRESO',taCreat,0,PRien) ;    // INUTILISE
    45102 : ParamTable('CPLUMIEREGENE',taCreat,0,PRien,17) ; // INUTILISE
    45103 : ParamTable('CPLUMIEREAUXI',taCreat,0,PRien,17) ; // INUTILISE
{$ENDIF} // IFNDEF EAGLCLIENT

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

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
  // Mode pack avancé ?
  VH^.OkModCPPackAvance := False ; // Par défaut non activé // MODIF PACK AVANCE
  // Gestion du multi-lingue ?
  VH^.OkModMultilingue  := False ; // Par défaut non activé // MODIF PACK AVANCE
  // RR dans tous les cas on considère que nous avons cette seria
  VH^.OkModAna := True ;
  VH^.OkModCPPackIFRS := False ;
      VH^.OkModCompta:=(sAcces[1]='X') ;
      {$IFDEF AMORTISSEMENT}
        VH^.OkModImmo:=(sAcces[2]='X') ;
      {$ELSE}
        VH^.OkModImmo:=FALSE ;
      {$ENDIF}
        If Not VH^.OldTeleTrans Then VH^.OkModEtebac:=TRUE ;
        // Mode S5 pack avance : positionne la série // MODIF PACK AVANCE
        if ( VH^.OkModCPPackAvance ) or ( V_PGI.VersionDemo ) then
        begin
            V_PGI.LaSerie := S7 ;
            // TEST Préférences dans agl pour style du menu :
            V_PGI.ChangeModeOutlook := True ;
        end ;

  // Si version démo --> Tout accessible mais limité
  
    if V_PGI.VersionDemo then
     BEGIN
{$IFDEF AMORTISSEMENT}
     V_PGI.NbColModuleButtons:=2 ; V_PGI.NbRowModuleButtons:=5 ;
     FMenuG.SetModules([9,11,13,16,14,12,20,8,17,18],[108,78,81,90,58,22,1,11,28,49]) ;
{$ELSE}
     V_PGI.NbColModuleButtons:=2 ; V_PGI.NbRowModuleButtons:=5 ;
     FMenuG.SetModules([9,11,13,16,14,12,8,17,18],[108,78,81,90,58,22,11,28,49]) ;
{$ENDIF}
     END Else
      // Si pas module compta --> pas les autres modules, seulement BAO
      if Not VH^.OkModCompta then
         BEGIN
         FMenuG.SetModules([17,18],[28,49]) ;
         V_PGI.NbColModuleButtons:=1 ;
         END else
         BEGIN // Nécessairement compta + éventuels autres modules
           if VH^.OkModBudget then
              BEGIN
              V_PGI.NbColModuleButtons:=2 ; V_PGI.NbRowModuleButtons:=5 ;
              if VH^.OkModImmo then FMenuG.SetModules([9,11,13,16,14,12,20,8,17,18],[108,78,81,90,58,22,1,11,28,49])
                               else FMenuG.SetModules([9,11,13,16,14,12,8,17,18],[108,78,81,90,58,22,11,28,49]) ;
              END else
              BEGIN
              if VH^.OkModImmo then
                 BEGIN
                 V_PGI.NbColModuleButtons:=2 ; V_PGI.NbRowModuleButtons:=5 ;
                 FMenuG.SetModules([9,11,13,16,14,20,8,17,18],[108,78,81,90,58,1,11,28,49]) ;
                 END else
                 BEGIN
                 V_PGI.NbColModuleButtons:=2 ; V_PGI.NbRowModuleButtons:=4 ;
                 FMenuG.SetModules([9,11,13,16,14,8,17,18],[108,78,81,90,58,11,28,49]) ;
                 END ;
              END ;
         END ;
{$IFDEF SCANGED}
if not VH^.OkModGed then OnGlobalAfterImportGED := nil;
{$ENDIF}
{$IFDEF TT}
VH^.OkModGed:= true ;
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
    8,9,11,13,14,16,17,18,52,53,54,55,56,96 : ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ; // FQ 10360
    12 : ChargeMenuPop(integer(hm3),FMenuG.DispatchX) ;
// BTY 07/05 Fiche 14754 Supprimer les anciennes éditions en cwas
    20,2 :
    begin
       ChargeMenuPop(integer(hm3),FMenuG.DispatchX) ;
       {$IFDEF EAGLCLIENT}
       if  (NumModule = 20) or (NumModule = 2) then
            FMenuG.RemoveGroup(20400,True) ;
       {$ENDIF}
    end;
// fin
    else ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
  END ;

  // Tripotage des menus
  TripotageMenu (NumModule);

  // Change les captions des menus Tiers / Décaissements
  if NumModule=11 then ChangeMenuDeca ;

  if FMenuG.PopupMenu=Nil then FMenuG.PopUpMenu:=ADDMenuPop(Nil,'','') ;

  {$IFNDEF EAGLCLIENT}
  FMDispS3.NumModule:=NumModule ;
  {$ENDIF}

  if (NumModule = 17)   then         // GP N° 5594
    FMenuG.RemoveItem(1272) ;

  // ajout me pour ca3
  if (NumModule = 96) then begin
  {$IFNDEF EAGLCLIENT}
  if GetParamSoc ('SO_ZGERETVA')= True then begin
{    if not ConnectDico then begin
      ConnectionDico;
      ConnectDico := TRUE;
      vb_SetUserInfoBd (InfoBdPgi);
    end;}
    end
  else {$ENDIF} FMenuG.RemoveGroup(7637,True) ;
  end;

  // ajout sb pour ca3
  if (NumModule = 14)   then begin
{    if GetParamSoc ('SO_ZGERETVA')= True then begin
      if not ConnectDico then begin
        ConnectionDico;
        ConnectDico := TRUE;
        vb_SetUserInfoBd (InfoBdPgi);
      end;
    end
    else }FMenuG.RemoveGroup(-7660,True) ;
  end;
END ;

function ChargeFavoris : boolean ;
begin
  if ctxPCL in V_PGI.PGIContexte
    then AddGroupFavoris( FMenuG , '54;2;55;96')
    else AddGroupFavoris( FMenuG , '12;20;8;17;18') ;
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
{$IFDEF CCS7}  // 13330
var
  i : integer;

  procedure RemoveMenu(M : TMenuItem);
  var
    j : Integer;
  begin
    for j := M.Count-1 downto 0 do begin
      if M.Items[j].Tag = iTag then begin
        M.Items[j].Visible := False;
        exit;
        end
      else if M.Items[j].Count >0 then RemoveMenu(M.Items[j]);
    end;
  end;
{$ENDIF}
begin
  if bGroup then FMenuG.RemoveGroup (iTag, True)
  else FMenuG.RemoveItem (iTag);
  stAExclure := stAExclure+IntToStr(iTag)+';';
{$IFDEF CCS7}
  if FMenuG.MMS7 <> nil then
    for i := 0 to FMenuG.MMS7.Items.Count-1 do begin
      if FMenuG.MMS7.Items[i].Tag = iTag then begin
        FMenuG.MMS7.Items[i].Visible := False;
        FMenuG.MM97.Controls[i].Visible := False;
        end
      else if FMenuG.MMS7.Items[i].Count > 0 then RemoveMenu(FMenuG.MMS7.Items[i]);
    end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/09/2003
Modifié le ... :   /  /
Description .. : - Suppression des menus
Suite ........ : - Constitution de la liste des tag à supprimer pour la gestion
Suite ........ : des favoris
Mots clefs ... :
*****************************************************************}
procedure TripotageMenuModePCL ( const TNumModule : array of integer; var stAExclure : string );
var i, NumModule : integer;
begin
  for i := 0 to High(TNumModule) do
  begin
    NumModule := TNumModule[i];
    case NumModule of
      { Outils comptables }
      53 :  begin
              { Suppression des éditions légales }
              RemoveFromMenu (7709, False, stAExclure);
            end;
      { Paramétrages }
      54 :  begin
              {$IFNDEF EAGLCLIENT}
              if GetParamSoc('SO_CPPCLSANSANA') then RemoveFromMenu (54600, True, stAExclure);
              {$ENDIF}
              { Réception ETEBAC }
              if (not VH^.OldTeleTrans) then RemoveFromMenu (1732, False, stAExclure);
            end;
      { Missions complémentaires }
      55 :  begin
              { Si Budget pas sérialisé, le menu Budget n'apparaît pas ! }
              if VH^.OkModBudget=False then RemoveFromMenu (55300, True, stAExclure);
              { Suppression analytique }
              {$IFNDEF EAGLCLIENT}
              if GetParamSoc('SO_CPPCLSANSANA') then RemoveFromMenu (7361, True, stAExclure);
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
              { Suppression du menu "contrôles" en Web Access }
              {$IFDEF EAGLCLIENT}
              RemoveFromMenu (56200, True , stAExclure);
              {$ENDIF}
            end;
      { Traitements annexes }
      96 :  begin
              { Si ICC pas sérialisé, le menu ICC n'apparaît pas }
              if VH^.OkModIcc=False then RemoveFromMenu (96100, True, stAExclure);
              { Si CRE pas sérialisé, le menu CRE n'apparaît pas }
              if VH^.OkModCre=False then RemoveFromMenu (96200, True, stAExclure);
              { Si Immo pas sérialisé, le menu Immo n'apparaît pas }
              if VH^.OkModImmo=False then RemoveFromMenu (20, True, stAExclure);
              { Si TVA non sérialisée, le menu TVA n'apparaît pas }
              if VH^.OkModTVA=False then RemoveFromMenu (7637, True, stAExclure)
              else
              begin
                RemoveFromMenu (7640, False, stAExclure);
                RemoveFromMenu ( 7648, False, stAExclure);
                RemoveFromMenu ( -51, False, stAExclure);
                RemoveFromMenu ( 7650, False, stAExclure);
                RemoveFromMenu ( 7651, False, stAExclure);
                RemoveFromMenu ( 7652, False, stAExclure);
              end;
              { Réception ETEBAC }
              if (not VH^.OldTeleTrans) then RemoveFromMenu ( 7772, False, stAExclure);
              {$IFNDEF EDICOMPTA}
              RemoveFromMenu ( 96, True, stAExclure);
              {$ENDIF}
            end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/10/2003
Modifié le ... : 09/10/2003
Description .. : - Suppression des menus Pack avancé
Suite ........ : - Constitution de la liste des tag à supprimer pour la gestion
Suite ........ : des favoris PGE
Mots clefs ... :
*****************************************************************}
procedure CacherMenuPackAvance ( const TNumModule : array of integer; var stAExclure : string );
var
  i, NumModule : integer;
  ePackAvance  : Boolean; {JP 09/06/05 : mise ne place du pack avancé en eAgl}
begin
  {$IFDEF EAGLCLIENT}
  {On part du principe, pour l'instant, que le pack avancé en CWAS est limité : on cache donc les menus
   comme si on n'était pas en pack avancé, sauf quelques uns qu'on active par un if not ePackAvance}
  ePackAvance := EstComptaPackAvance;
  {$ELSE}
  {Si pack avancé sérialisé, on sort...}
  if EstComptaPackAvance then Exit ;
  ePackAvance := False;
  {$ENDIF}

  for i := 0 to High(TNumModule) do
  begin
    NumModule := TNumModule[i];
    case NumModule of
      { Ecritures }
      9 : begin
          // SBO : Report V620 : sinon les menu prévision et révision apparaissent en mode S5 sans pack avancé !
          RemoveFromMenu( 7300, True, stAExclure);  //Ecritures de Prévision
          RemoveFromMenu(76460, True, stAExclure); //Ecritures de Révision
          {
          RemoveFromMenu(7670, False, stAExclure) ; // Révision : Brouillard
          RemoveFromMenu(7682, False, stAExclure) ; // Révision :  Mode révision
          }

          {JP 08/06/05 : Migration eAgl / ePack avancé}
          if not ePackAvance then begin
            RemoveFromMenu (7371, False, stAExclure); // Ecritures analytiques : Modifications en série sur tables libres
            RemoveFromMenu (7379, False, stAExclure); // Ecritures analytiques : Transfert intersection ;
            //SG6 19/01/05 FQ 15270 Modification tables libres
            RemoveFromMenu (7260, False, stAExclure); // Ecriture courante :  Modification en série sur table libre
            RemoveFromMenu (7282, False, stAExclure); // Ecritures de simulation : Modifications en série sur tables libres
            RemoveFromMenu (7305, False, stAExclure); // Ecritures de situation : Modifications en série sur tables libres
            RemoveFromMenu(7317, False, stAExclure) ; // Ecritures de prévision : Modifications en série sur tables libres
          end;
        end ;
      { BUDGETS }
      12 :  begin
            // Structure
            RemoveFromMenu (-15118, False, stAExclure); // Modifications en série sur tables libres
            RemoveFromMenu (15161, False, stAExclure); // Fermeture des Budgets
            RemoveFromMenu (15163, False, stAExclure); // Réouverture des Budgets
            // Editions
            RemoveFromMenu (15340, False, stAExclure); // Etats multi-budgets
            RemoveFromMenu (15350, False, stAExclure); // Etats libres
            // Paramètres
            RemoveFromMenu (1498, False, stAExclure); // Catégorie budgetaires
            RemoveFromMenu (1504, False, stAExclure); // Automate de création des comptes à partir des tables libres
            RemoveFromMenu (1524, False, stAExclure); // Automate de création des Sections à partir des tables libres
            RemoveFromMenu (1529, False, stAExclure); // Automate de création des sections à partir des catégories budgétaires
            end ;
      { Editions }
      13 :  begin
              {JP 08/06/05 : Migration eAgl / ePack avancé : les tables libres ont été migrées}
              if not ePackAvance then begin
                // Editions sur tables libres ecritures
                RemoveFromMenu (7419, False, stAExclure); // Grand livre
                RemoveFromMenu (7429, False, stAExclure); // Balance

                //SG6 19/01/05 FQ 15267 Editions multi dossier
                RemoveFromMenu (8614, False, stAExclure); // Balance sur tables libres
                RemoveFromMenu (8624, False, stAExclure); // GrandLivre sur tables libres
              end;
              // Autres
              RemoveFromMenu (7490, False, stAExclure); // Récapitulatifs des éditions
            end;
      { Traitement }
      16 :  begin
            RemoveFromMenu (7719, False, stAExclure);   // Comptes : Affectation sur tables libres écritures
            RemoveFromMenu (77120, False, stAExclure);  // Clotures / validations -> Etats des éditions légales
            RemoveFromMenu (-7625, True, stAExclure);   // Analyse de trésorie
            end;
      { Structures et paramètres }
      17 :  begin
            RemoveFromMenu (-7116, False, stAExclure); // Modifications en série
            RemoveFromMenu (1155, False, stAExclure); // Langues
            RemoveFromMenu (1195, False, stAExclure); // Personnalisation des zones libres des mvts
            RemoveFromMenu (1197, False, stAExclure); // Recopie des tables libres
            RemoveFromMenu (1526, False, stAExclure); // Génération des tables libres à partir des sous-sections
            {JP 08/06/05 : Migration eAgl / ePack avancé}
            if not ePackAvance then
              RemoveFromMenu (1455, False, stAExclure); // Compta analytique : Clés de répartition
            end ;
      { Outils et administration }
      18 :  begin
            if not VH^.OkModMultilingue then RemoveFromMenu (3100, True, stAExclure); // Menu installation // 14693
            RemoveFromMenu (3227, False, stAExclure); // Menu Gestion des analytiques //SG6 08/12/2004
            end;
    end;
  end;
end;

function IsOracle : Boolean;
begin
  Result := (V_PGI.Driver=dbOracle7) or (V_PGI.Driver=dbOracle8);
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

// Plus de notion de pack avancé en mode PGE...
// Les non-compatibilité du à l'eAGL sont gérés ci dessous.
          // Gestion du pack avancé (dans ENT1)
          // CacherMenuPackAvance(TNumModule, stAExclure);

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
             RemoveFromMenu (7268, False, stAExclure) ; // Brouillard
             RemoveFromMenu (7289, False, stAExclure) ; // Brouillard
             RemoveFromMenu (7381, False, stAExclure) ; // Ré-imputations
             RemoveFromMenu (7385, False, stAExclure) ; // Brouillard
             RemoveFromMenu(7670, False, stAExclure) ; // Révision : Brouillard
             RemoveFromMenu(7682, False, stAExclure) ; // Révision :  Mode révision
           {$ENDIF}
           // Ecriture IFRS
           if not EstComptaIFRS then
             RemoveFromMenu (-9200, True, stAExclure) ; // Saisie IFRS
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
             RemoveFromMenu (7598,   False, stAExclure) ;   // Estimation du coût des tiers

             // Editions de suivi :
             RemoveFromMenu (7535,   False, stAExclure) ;   // Justif de solde
             RemoveFromMenu (7550,   False, stAExclure) ;   // GL Agé
             RemoveFromMenu (7559,   False, stAExclure) ;   // GL Ventilé
             RemoveFromMenu (-7230,  True, stAExclure) ;   // Tiers Payeurs
             RemoveFromMenu (-52700, True, stAExclure) ;   // Autres éditions

             // --> Emissions
             RemoveFromMenu (7599,   False, stAExclure) ;  // Bordereau d'accompagnement
             RemoveFromMenu (-7586,  False, stAExclure) ;  // Export CFONB
             RemoveFromMenu (-27,    False, stAExclure) ;  // Ré-éditions de traites et BOR
           {$ENDIF}
           // Cacher le module d'écart de conversion du lettrage
           RemoveFromMenu (7524, False, stAExclure) ; // Lettrage --> Ecart de conversion
           end ;

      // ===============
      // === BUDGETS ===
      // ===============
      12 : begin
           {$IFDEF EAGLCLIENT}
             // Structure
             RemoveFromMenu (-15118, False, stAExclure); // Modifications en série sur tables libres
             RemoveFromMenu (15161, False, stAExclure); // Fermeture des Budgets
             RemoveFromMenu (15163, False, stAExclure); // Réouverture des Budgets
             // Editions
             RemoveFromMenu (15340, False, stAExclure); // Etats multi-budgets
             RemoveFromMenu (15350, False, stAExclure); // Etats libres
             // Paramètres
            RemoveFromMenu (-4, False,  stAExclure) ; //FP FQ16051 Cacher les éditions d'écarts budgetaires
            RemoveFromMenu (-5, False,  stAExclure) ; //FP FQ16052 Cacher les éditions des brouillards budgetaires
             RemoveFromMenu (1498, False, stAExclure); // Catégorie budgetaires
             RemoveFromMenu (1504, False, stAExclure); // Automate de création des comptes à partir des tables libres
             RemoveFromMenu (1524, False, stAExclure); // Automate de création des Sections à partir des tables libres
             RemoveFromMenu (1529, False, stAExclure); // Automate de création des sections à partir des catégories budgétaires
           {$ENDIF EAGLCLIENT}
           end ;

      // ================
      // === EDITIONS ===
      // ================
      13 : begin
           {$IFDEF EAGLCLIENT}
             // Journaux
             RemoveFromMenu (7409, False, stAExclure);   // Périodique par général
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
           else if VH^.PaysLocalisation=CodeISOFR then
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
             RemoveFromMenu (7631, False, stAExclure) ; // Contrôle de caisse
             // Validations
             RemoveFromMenu (7709, False, stAExclure) ; // Editions légales
             // Clôtures
             RemoveFromMenu (7757, False, stAExclure) ;
             RemoveFromMenu (7760, False, stAExclure) ;
             RemoveFromMenu (7766, False, stAExclure) ;
             RemoveFromMenu (7769, False, stAExclure) ;
             RemoveFromMenu (7683, False, stAExclure) ;
             // Relevés
             RemoveFromMenu (7772, False, stAExclure) ;
           {$ENDIF}
             {JP 17/08/05 : FQ 15855 : suppression du menu dans la SocRef 704 du 25/08/05
             RemoveFromMenu(7607, False, stAExclure ) ; // Dépointage}
           end;

      // ================================
      // === Structures et paramètres ===
      // ================================
      17 : begin
           {$IFDEF EAGLCLIENT}
             // Paramètres //RR 02/06/2003
             RemoveFromMenu(1315, False, stAExclure ) ; // Plans de référence
             RemoveFromMenu(-9, False, stAExclure ) ;   // Comptes de correspondance
             RemoveFromMenu(-10, False, stAExclure ) ;  // Plans de ruptures des éditions
             RemoveFromMenu(-12, False, stAExclure ) ;  // Etapes de réglement
             // Banques
             RemoveFromMenu(1730, False, stAExclure ) ; // Codes résidents étrangers
             RemoveFromMenu(1272, False, stAExclure ) ; // Paramètrage
             // Documents
             RemoveFromMenu (1225, True, stAExclure) ;
             RemoveFromMenu (-1280, False, stAExclure) ; // Edition paramétrables NE PAS SUPPRIMER CETTE LIGNE
             // Non compatible eAGL...
//             RemoveFromMenu (-7116, False, stAExclure); // Modifications en série
//             RemoveFromMenu (1155, False, stAExclure); // Langues
//             RemoveFromMenu (1195, False, stAExclure); // Personnalisation des zones libres des mvts
             RemoveFromMenu (1197, False, stAExclure); // Recopie des tables libres
             RemoveFromMenu (1526, False, stAExclure); // Génération des tables libres à partir des sous-sections
           {$ELSE}
             RemoveFromMenu (1266, False, stAExclure) ; // Paramétrage lettre prélèvement
             RemoveFromMenu (1269, False, stAExclure) ; // Paramétrage lettre virement
           {$ENDIF}

           RemoveFromMenu (-1426, False, stAExclure); // Compte de regroupement
           if not EstComptaTreso then
             begin
             RemoveFromMenu (1706, False, stAExclure); // Agences bancaires
             RemoveFromMenu (1707, False, stAExclure); // Comptes bancaires
             end;
           if not EstMultiSoc then
             RemoveFromMenu (1755, False, stAExclure) ;

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
             RemoveFromMenu (3185, False, stAExclure); // Suivi activité
             // Outils
             RemoveFromMenu (3210, False, stAExclure);
             RemoveFromMenu (3215, False, stAExclure);
             RemoveFromMenu (3220, False, stAExclure);
             RemoveFromMenu (3221, False, stAExclure);
             RemoveFromMenu (3227, False, stAExclure); //Changement du mode analytique
           {$ENDIF}
           end;

      // =====================
      // === Amortissement ===
      // =====================
      20 : Begin
           {$IFDEF EAGLCLIENT}
             FMenuG.RemoveGroup(20400,True);
           {$ENDIF}
           end ;
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
  RemoveFromMenu (-20260, False, stAExclure);
  { Etat comparatif des bases TP }
  RemoveFromMenu (20254, False, stAExclure);
  { Etiquettes }
  RemoveFromMenu (20276, False, stAExclure);
  { Fiche Immobilisation }
  RemoveFromMenu (20278, False, stAExclure);
  { Clôture }
  // RemoveFromMenu (2416 , False, stAExclure);
  { Import/Export }
  {$IFDEF EAGLCLIENT}
  RemoveFromMenu (20511, False, stAExclure);
  RemoveFromMenu (20512, False, stAExclure);
  {$ENDIF}
end;

procedure TripotageMenu (NumModule : integer);
var stAExclure : string;
begin
  if ctxPCL in V_PGI.PGIContexte
    then TripotageMenuModePCL ( [NumModule] , stAExclure )
    else TripotageMenuModePGE ( [NumModule] , stAExclure ) ;
  if (NumModule = 20) or (NumModule = 2 ) then TripotageMenuAmortissement (stAExclure);
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

{$IFDEF EAGLCLIENT}
{$IFDEF PORTAGEBUDGET}
if Not Assigned(ProcZoomBudGen)  then ProcZoomBudGen  :=FicheBudgene;
if Not Assigned(ProcZoomBudSec)  then ProcZoomBudSec  :=FicheBudsect;
{$ENDIF}
// If Not Assigned(ProcZoomNatCpt)  Then ProcZoomNatCpt  :=FicheNatCpte; // A FAIRE : Nécessite NATCPTE_TOM.PAS
{$ELSE}
if Not Assigned(ProcZoomBudGen)  then ProcZoomBudGen  :=FicheBudgene ;
if Not Assigned(ProcZoomBudSec)  then ProcZoomBudSec  :=FicheBudsect ;
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

      sDom:='05990011' ;
      VH^.SerProdCompta       := '05010060';
      VH^.SerProdBudget       := '05020060';
      VH^.SerProdImmo         := '05040060';
      VH^.SerProdEtebac       := '05030011';
      VH^.SerProdMultilingue  := '00088060'; // MODIF MULTILINGUE
      VH^.SerProdPackIFRS     := '00232060'; // MODIF IFRS
      {$IFDEF AMORTISSEMENT}
      FMenuG.SetSeria(sDom,[VH^.SerProdCompta,VH^.SerProdImmo],
                            ['Business Suite','Amortissement']) ;
      {$ELSE}
      FMenuG.SetSeria(sDom,[VH^.SerProdCompta],
                            ['Business Suite']) ;
      {$ENDIF}

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
    {$IFNDEF EAGLCLIENT}
    FMenuG.Outlook.OnChangeActiveGroup:=FMDispS3.AfterChangeGroup;
    {$ENDIF}
    {$IFDEF SCANGED}
    OnGlobalAfterImportGED := MyAfterImport;
    {$ENDIF}
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 19/01/2005
Modifié le ... :   /  /    
Description .. : - LG - 19/01/2005 - utilisation de la fct FicheTiersZoom. On 
Suite ........ : recup le range pour ouvrir la fiche des tiers sur l'auxi courant
Mots clefs ... :
*****************************************************************}
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
  case Num of
    // general
    1,7112 : FicheGene(Nil,'',LeQuel,Action,0);
    // tiers
    2,7145 : FicheTiersZoom(Nil,'',LeQuel,Action,1,Range) ;
    // journal
    4,7211 : FicheJournal(Nil,'',Lequel,Action,0) ;

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
                taModif : YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT');
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
    end ;
end ;

{$IFNDEF EAGLCLIENT}
procedure TFMDispS3.AfterChangeGroup(Sender : TObject);
begin
  if ctxPCL in V_PGI.PGIContexte then
  begin
    if (NumModule=55) then
    begin
      { FQ 15736 - Si dernier groupe de la liste, alors c'est le budget ==> popup particulier }
      if (fMenuG.Outlook.Groups.Count = fMenuG.Outlook.ActiveGroup+1) then
       ChargeMenuPop(integer(hm3),FMenuG.DispatchX)
      else ChargeMenuPop(integer(hm2),FMenuG.DispatchX);
    end
    else ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ;
  end;
end;

procedure TFMDispS3.ExercClick(Sender: TObject);
begin
  // CA - 27/11/2001
  ParamSociete(False,'','SCO_DATESDIVERS','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,0) ;
end;

procedure TFMDispS3.FMDispS3Create(Sender: TObject);
begin
  PGIAppAlone:=True ;
  CreatePGIApp ;
  SeriaMessOk := False;
end;
{$ENDIF}

Procedure InitLaVariablePGI;
Begin
  {Version}
  Apalatys:='CEGID' ;
  HalSocIni:='CEGIDPGI.INI' ;
  Copyright:='© Copyright ' + Apalatys ;
  V_PGI.CodeProduit  := '001';
{$IFDEF CCS7}
  RenseignelaSerie(ExeCCS7) ;
{$ELSE}
  RenseignelaSerie(ExeCCS3) ;
{$ENDIF}

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
  V_PGI.EuroCertifiee:=True;
{$ELSE}
  V_PGI.EuroCertifiee:=False;
{$ENDIF}

  ChargeXuelib ;

  {Série}
{$IFDEF CCS7}
  V_PGI.LaSerie:=S7 ;
{$ELSE}
  V_PGI.LaSerie:=S3 ;
{$ENDIF}

  // 13584
  if ctxPCL in V_PGI.PGIContexte then V_PGI.PortailWeb := 'http://experts.cegid.fr/home.asp'
                                 else V_PGI.PortailWeb := 'http://utilisateurs.cegid.fr';

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
{$IFDEF CCS7}
  RenseignelaSerie(ExeCCS7) ;
{$ELSE}
  RenseignelaSerie(ExeCCS3) ;
{$ENDIF}

  // Initialisé à Vrai par défaut !
  V_PGI.EnableTableToView := False ;

End;


Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;
  RegisterHalleyWindow ;
end.

 













