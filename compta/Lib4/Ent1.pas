{***********UNITE*************************************************
Auteur  ...... : ROHAULT Régis
Créé le ...... : 19/07/2005
Modifié le ... : 19/07/2005
Description .. : Attention, les séries S3 et S7 ne devront plus exister.
Suite ........ : S5 est la seule série utilisable.
Suite ........ :
Suite ........ : Buisiness  Suite et Buisiness  Place sont tout 2 des S5.
Mots clefs ... : S3;S5;S7;PLACE;SUITE;SERIE
*****************************************************************}
unit Ent1;

interface

uses
     Forms,
     {$IFNDEF EAGLSERVER}
     buttons,
     Graphics,
     Grids,
     Windows,
     Menus,
     Hrichedt,
     HTB97,
     HMsgBox,
     {$ENDIF}
     IniFiles,dialogs,SysUtils,TabNotBk,Controls,
     Classes,
     ComCtrls,
     M3FP,
{$IFDEF EAGLCLIENT}
     UtileAGL,
     MenuOlx, // GetLastNumCLick
{$ELSE}
     DBGrids,
     DBCtrls,
     DB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     MajTable,
     HDB,
     {$IFNDEF EAGLSERVER}
     MenuOlg, // GetLastNumCLick
     {$ENDIF}
{$ENDIF}
     HCtrls,
     stdctrls,
     ExtCtrls,
     HEnt1,
     registry,
     WinProcs,
     shellAPI,
     Hstatus,
     HCrc,
     LicUtil,
     Messages,
     UTOB,
     CBPPath,
     {$IFNDEF NOVH}
     {$IFNDEF HVCL}
     EntPGI,
//  BBI Web services
     {$ENDIF HVCL}
     {$ENDIF NOVH}
     {$IFNDEF EAGLSERVER}
     {$IFNDEF EAGLCLIENT}
      uHListe, //,seriapgi
     {$ENDIF EAGLCLIENT}
     {$ENDIF EAGLSERVER}
//  BBI Fin Web services
     {$IFDEF EAGLSERVER}
      eSession ,
      uWa ,
      {$ENDIF}
{$IFDEF COMPTA}
     galOutil,
{$ENDIF}
     FiltresDonnees, // MB : Nécessite AGL 20.21 //
     // Rajouter la condition KPMG pour continuer
     // utiliser la nouvelle fonctionnalités
    ed_tools
    {$IFNDEF EAGLSERVER}
    {$IFDEF EAGLCLIENT}
    {$ELSE}
      , Web
    {$ENDIF}
    {$ENDIF}
    // MB Le 29/11/2007
    {$IFNDEF EAGLSERVER}
    , PGIExec
    {$ENDIF EAGLSERVER}
    ,UentCommun
    ,UHTTP
    ;

Const  WB_MOVEWORDLEFT    = 4;
       WB_MOVEWORDRIGHT   = 5;
       EM_FINDWORDBREAK   = WM_USER + 76;

Const GMM : Integer = 1 ;
(* GP Code Spécif :
51199 : Escompte fournisseur   Passe en standard
51198 : Modid entete pièce     Passe en standard
51197 : Nouveau module relance Passe en standard
51196 : Affectation automatique TL0 tiers sur TL0 ecr  Passe en standard ???  (mais pas sur la saisie)
51195 : Lettre virements pour CCMP  Passe en standard
51194 : Consultation des écritures en PGE : ancienne version non optimisée
51193 : Accès à la saisie en mode libre    Passe en standard
51192 : Lettre prélèvements
51191 : Accès à l'ancienne version Pointage
51190 : Accès aux anciennes éditions de pointage
51189 : Accès à l'ancienne version extourne
51188 : Stop écart de conversion
51187 : GetCumuls optimisé
51186 : Onglet avancès sur éditions standards
51185 : Ancienne version de getcumuls non compatible EAGL
51500 : CCS7 en mode conso gestion
51501 : creation d'un fichier de sauvegarde en saisie de bordereau
51502 : Gestion des standard niveau Cegid
51503 : Accès à l' ancienne version de la consultation des écritures
51204 : Sauvegarde du folio en cours de saisie
51205 : Autorise tous les caractères pour G_GENERAL ou T_AUXILIAIRE
51206 : Encaissement / Décaissement en masse (DIRECT NRJ)
51207 : Pas de sauvegarde dossier format TRA au lancement de la PURGE
51208 : Saisie paramétrable en PCL
51209 : ancien reclassement
51210 : validation d'écritures simulées avec une REFGESCOM
51211 : activation GED en entreprise
51212 : Paramétrage des masques en saisie
51213 : Activation du nouveau pointage
51214 : Choix de la devise dans les états de synthèse
51215 : Enregistrement spécifique des TOb en saisie (sans les Blob) pour pb CWAS
51310 : Specif CERIC
51320 : Specif TESSI TVA : possibilité de réimputation sur pièce GESCOM + onglet avancé (CPREIMPUT_TOF.PAS)
*)

{================= Description LAVARIABLE / Parasmoc =========}
Type TExoDate = RECORD
                Code    : String[3] ;
                Deb,Fin : TDateTime ;
                DateButoir,DateButoirRub,DateButoirBud,DateButoirBudgete : TDateTime ;
                NombrePeriode : Integer ;
                EtatCpta : String ;
                END ;

Type TTabExo = Array[1..5] Of TExoDate ;

Type UneTaxe = RECORD
              Regime : String[3] ;
              Code : String[3] ;
              TauxACH,TauxVTE : Double ;
              CpteACH,CpteVTE : String[17] ;
              EncaisACH,EncaisVTE : String[17] ;
              END ;
     LesTaxe = Array[1..100] of UneTaxe ;

type  TCPRoleCompta = (rcCollaborateur, rcReviseur, rcSuperviseur, rcExpert);

type  TCPRevision = record
                      Plan : string;
                      DossierPretSupervise : Boolean;
                      DossierSupervise     : Boolean;
                    {$IFNDEF EAGLSERVER}
                      BStatutRevision      : TToolBarButton97;
                    {$ENDIF}
                    end ;

Const
      MaxEche = 50 ;
      MaxCatBud = 10 ;
      MaxTableLibre = 9 ;
      NBDECIMALTAUX = 6; {JP 25/05/05 : Nombre de décimale sur les taux de changes des devises}

Const JalPurge : String = '---' ;
      CharJal : String = 'S' ;


Const CS3 : Boolean = FALSE ;

Type TSectRetri=(srOk,srNonStruct,srPasEnchainement,srBadArg) ;
Type TStructure = (tstAux) ;
Type TTableLibreCompta = (tlGene,tlAux,tlSect,tlCptBud,tlSectBud,tlEcrGen,tlEcrAna,tlEcrBud,tlImmo) ;
Type TParamTableLibre=Array[tlGene..tlImmo,1..10] Of Boolean ;

Const { Concepts }
      {Comptabilité}
      ccGenCreat=11  ; ccGenModif=12    ; ccAuxCreat=13       ; ccAuxModif=14 ;
      ccSecCreat=15  ; ccSecModif=16    ; ccJalCreat=17       ; ccJalModif=18 ;
      ccSaisSolde=20 ; ccSaisProrata=21 ; ccSaisCreatGuide=22 ;
      ccSaisCreatVentil=23 ; ccSaisEcritures=24 ;
      ccSaisMvtBqe = 25; {JP 13/02/07 : Gestion du concept pour les mouvements bancaires}
      ccimport = 26; ccexport = 27; // ajout me 17/07/2007
      {Gestion Commerciale}
      gcArtCreat=61;
      gcArtModif=62;
      gcProspectCreat=63;
      gcForceRisque=64;
      gcCliCreat=65;
      gcSaisModifRepres=66;
      gcAdminECommerce=67;
      gcNumCBComplet=68;
      gcOutreVisaMulti=69;
      gcConsultStock=70;
      gcFouCreat = 71;
      gcPieceDelete = 72;
      gcTransfertSusPro=73;
      gcTauxNegocie=74;
      GcCatalogue=75;
      {Immo PGI}
      ciImoCreat=80 ; ciImoModif=81 ; ciCumAnte=82 ; ciModifDroitAide=83 ; ciAideSaiCreat=84 ; ciAideSaiModif=85 ;
      ciModifMultiDossier=86; ciUseMultiDossier=87;
      {Trésorerie}
      ciCreSupFlux = 53; ciValidationBO = 55;
      {Gestion Commerciale - 2}
      gcSavMajParc=100;
      gcSavMajArticleParc=101;
      gcSavMajVersion=102;
      gcSavMiseEsHs=103;
      gcSavReaffectEtMaj=104;
      gcArtModifCreat=105;
      gcProspectModif=510;
      gcCliModif=511;
      gcImportModifTiers=114; //modification des tiers sur import des suspects ;
      gcVisaCiblage=115;  // attribution du visa sur le ciblage
      { visu tiers 360 }
      gcFouModif=518;
      gc360Prospect=118;
      gc360Client=119;

      {Affaire}
      ga360Affaire = 118;

Type TStructureUniq = RECORD
                      CodeSp : String ;
                      QuelCarac : String ;
                      Debu,Long : Integer ;
                      End ;

Type TInfoCpta = RECORD
             Lg : Byte ;
             Cb : Char ;
             Chantier  : Boolean ;
             Structure : Boolean ;
             Attente      : String ; // LG 25/06/2007 - correction d'une fuite memmoire ds ChargeVHImmo
             AxGenAttente : String ; // LG 25/06/2007 - correction d'une fuite memmoire ds ChargeVHImmo
             UneStruc : TStructureUniq ;
             END ;


Type ttTypePiece  = (tpReel,tpSim,tpPrev,tpSitu,tpRevi,tpCloture,tpIfrs) ;  // Modif IFRS 05/05/2004
Type SetttTypePiece = Set Of ttTypePiece ;


Type TZoomTable = (tzGeneral,tzGCollectif,tzGNonCollectif,TzGCollClient,tzGCollFourn,
                   tzGCollDivers,tzGCollSalarie,tzGventilable,tzGlettrable,tzGCollToutDebit,tzGCollToutCredit,
                   tzGToutDebit,tzGToutCredit,tzGVentil1,tzGVentil2,tzGVentil3,tzGVentil4,tzGVentil5,
                   tzGpointable,tzGbanque,tzGcaisse,tzGTID,tzGTIC,tzGTIDTIC,tzGcharge,tzGBilan,
                   tzGproduit,tzGimmo,tzGDivers,tzGextra,tzGLettColl,tzGBQCaiss,tzGBQCaissCli,
                   tzGBQCaissFou,tzGEncais,tzGDecais,tzGNLettColl,tzGEffet,

                   tzTiers,tzTclient,tzTfourn,tzTsalarie,tzTdivers,tzTDebiteur,tzTCrediteur,
                   tzTlettrable,tzTToutDebit,tzTToutcredit,tzTReleve,

                   tzSection,tzSection2,tzSection3,tzSection4,tzSection5,

                   tzBudgen,tzBudSec1,tzBudSec2,tzBudSec3,tzBudSec4,tzBudSec5,

                   tzBudJal,tzBudgenAtt,tzBudSecAtt1,tzBudSecAtt2,tzBudSecAtt3,tzBudSecAtt4,tzBudSecAtt5,

                   tzJournal,tzJvente,tzJachat,tzJbanque,tzJcaisse,tzJOD,tzJAN,tzJCloture,tzJEcartChange,
                   tzJAna,tzJAna1,tzJAna2,tzJAna3,tzJAna4,tzJAna5,

                   tzArticle,tzArticlePlus,tzNomenclature,
                   tzImmo,
                   tzCorrespGene1,tzCorrespGene2,tzCorrespAuxi1,tzCorrespAuxi2,tzCorrespBud1,tzCorrespBud2,
                   tzCorrespSec11,tzCorrespSec12,tzCorrespSec21,tzCorrespSec22,tzCorrespSec31,tzCorrespSec32,
                   tzCorrespSec41,tzCorrespSec42,tzCorrespSec51,tzCorrespSec52,
                   tzNatGene0,tzNatGene1,tzNatGene2,tzNatGene3,tzNatGene4,
                   tzNatGene5,tzNatGene6,tzNatGene7,tzNatGene8,tzNatGene9,
                   tzNatTiers0,tzNatTiers1,tzNatTiers2,tzNatTiers3,tzNatTiers4,
                   tzNatTiers5,tzNatTiers6,tzNatTiers7,tzNatTiers8,tzNatTiers9,
                   tzNatSect0,tzNatSect1,tzNatSect2,tzNatSect3,tzNatSect4,
                   tzNatSect5,tzNatSect6,tzNatSect7,tzNatSect8,tzNatSect9,
                   tzNatBud0,tzNatBud1,tzNatBud2,tzNatBud3,tzNatBud4,
                   tzNatBud5,tzNatBud6,tzNatBud7,tzNatBud8,tzNatBud9,
                   tzNatBudS0,tzNatBudS1,tzNatBudS2,tzNatBudS3,tzNatBudS4,
                   tzNatBudS5,tzNatBudS6,tzNatBudS7,tzNatBudS8,tzNatBudS9,
                   tzNatEcrE0,tzNatEcrE1,tzNatEcrE2,tzNatEcrE3,
                   tzNatEcrA0,tzNatEcrA1,tzNatEcrA2,tzNatEcrA3,
                   tzNatEcrU0,tzNatEcrU1,tzNatEcrU2,tzNatEcrU3,
                   tzAxe1SP1,tzAxe1SP2,tzAxe1SP3,tzAxe1SP4,tzAxe1SP5,tzAxe1SP6,
                   tzAxe2SP1,tzAxe2SP2,tzAxe2SP3,tzAxe2SP4,tzAxe2SP5,tzAxe2SP6,
                   tzAxe3SP1,tzAxe3SP2,tzAxe3SP3,tzAxe3SP4,tzAxe3SP5,tzAxe3SP6,
                   tzAxe4SP1,tzAxe4SP2,tzAxe4SP3,tzAxe4SP4,tzAxe4SP5,tzAxe4SP6,
                   tzAxe5SP1,tzAxe5SP2,tzAxe5SP3,tzAxe5SP4,tzAxe5SP5,tzAxe5SP6,
                   tzRubCPTA,tzRubBUDG,tzRubBUDS,tzRubBUDSG,tzRubBUDGS,
                   tzNatImmo0,tzNatImmo1,tzNatImmo2,tzNatImmo3,tzNatImmo4,
                   tzNatImmo5,tzNatImmo6,tzNatImmo7,tzNatImmo8,tzNatImmo9,
                   tzTPayeur,tzTPayeurCLI,tzTPayeurFOU,
                   tzRien
                   ) ;

Type TfourCpt = RECORD
                Deb : String[17] ;
                Fin : String[17] ;
                END ;

Type TFBil = Array[1..5] Of TFourCpt ;

Type TFCha = Array[1..5] Of TFourCpt ;

Type TFPro = Array[1..5] Of TFourCpt ;

Type TFExt = Array[1..2] Of TFourCpt ;
Type TLGTABLELIBRE = Array[1..MaxTableLibre,1..10] Of Integer ;
Type TNIVTABLELIBRE = Array[1..MaxTableLibre,1..10] Of String[5] ;

Type TUneCatBud = Record
                  fb : TFichierBase ;
                  Code,Lib : String ;
                  SurJal  : Array[1..MaxSousPlan] Of String ;
                  SurSect : Array[1..MaxSousPlan] Of String ;
                  End ;
Type TCatBud = Array[1..MaxCatBud] Of TUneCatBud ;

Type TSousPlanCat = Array[1..MaxSousPlan] Of TSousPlan ;

Type TMPPop = Record
              MPGenPop,MPAuxPop,MPJalPop,MPExoPop : String ;
              MPDatePop : tDAteTime ;
              MPNumPop,MPNumLPop,MPNumEPop : Integer ;
              End ;

Type TProfilUser = Record
                   Compte1,Compte2,Exclusion1,Exclusion2,TableLibre1,TableLibre2 : String ;
                   Etablissement,Domaine : String ;
                   ForceEtab,ForceDomaine : Boolean ;
                   ProfilOk : Boolean ; // TRUE SI PROFIL TROUVE DANS CHARGEPROFIL
                   Depot,PCPVente,PCPAchat : string; //JTR Dépôt par défaut et PCP
                   ForceDepot : boolean; // JTR - Force le dépot
                   responsable : string;
                   ForceResponsable : boolean;
                   End ;

Type tProfilTraitement=(prClient,prFournisseur,prEtablissement,prAucun) ;

Type tProfil = Array[prClient..prEtablissement] Of tProfilUser ;

Type tSauveSA_SaisieEff= Array[0..15] Of ShortInt ;

Type tCCMP = Record
             LotCli,LotFou : Boolean ;
             End ;

Type LaVariableHalley = RECORD
     TpfDehors,TpfApres,PasParSoc,JalLookUp,EtabLookUp : Boolean ;
{$IFDEF SPEC302}
     TabTVA,TabTPF : LesTaxe ;
{$ELSE}
     LaTOBTVA : TOB ;
     OBImmo   : TOB ;
     MPACC    : TStrings ;
{$ENDIF}
     OuiTvaEnc,OuiTP,AttribRIBAuto : boolean ;
     ZSaisieEche,ZSaisieAnal,ZGereAnal,ZAutoSave,ZActivePFU,ZFolioTempsReel,ZSauveFolioLocal : Boolean ;
     ModeSilencieux,CPPCLSansAux,CPPCLSansAna,CPAnaAvancee : boolean ;
     CptesGeneCrees,CpteAuxiCrees : String ;
     CPControleDoublon,LienPublifi,CPDateObli,CPCodeAuxiOnly,CPLibreAnaObli,CPAppelOLE : boolean ;
     CPChampDoublon,CPStatusBarre,CPLienGamme : String ;
     CPPointageSx  : string; // ajout me 02/07/2004
     CPDaterecep   : TDateTime;
     NumCodeBase   : Array[1..255] of String[3] ;
     Cpta : Array[fbAxe1..fbAux] of TInfoCpta ;
     VerifCorresp : Array[1..10,1..2] of Boolean ;
     EnCours,Suivant,Precedent,Entree,ExoV8,ExoEdit,CPExoRef : TExoDate ;
     ExoClo : TTabExo ;
     DateCloturePer,DateCloturePro,Date1,Date2,Date3,DateRevision,CPLastSaisie : TDateTime ;
     NbjEcrAV,NbjecrAP,NbjEchAV,NbjEchAP : integer ;
     TvaEncSociete,EtablisDefaut,JalCtrlBud,{JalEcartEuro,RegleEquilSais,}JalVTP,JalATP,JalRepBalAN : String[3] ;
     CPRefLettrage,CPRefPointage : String[3] ;
     MontantNegatif,EtablisCpta,TenueEuro : Boolean ;
     DupSectBud,EquilAnaODA,AutoExtraSolde : Boolean ;
     JalAutorises,MonnaiesIn,LibDevisePivot,LibDeviseFongible,SymboleFongible : String ;
     GrpMontantMin,GrpMontantMax,TauxCoutTiers,MaxDebitPivot,MaxCreditPivot : double ;
     Pays : String[3] ;
     PaysLocalisation : String[3] ; //XVI 24/02/2005
     OuvreBil{,EccEuroDebit,EccEuroCredit} : String[17] ;
     //LG* 27/03/2002 suppression suite au modif du tobm
     //DescriEcr,DescriAna,DescriGui,DescriAbo,DescriSais,DescriBud : TStrings ;
     ListeNatJal,CollCliEnc,CollFouEnc,DefCatTVA,DefCatTPF : String ;
     ArrondiLigne               : Boolean ;
     DefautCli : String[17] ;
     DefautFou : String[17] ;
     DefautSal : String[17] ;
     TiersDefSAl,TiersDefDiv,TiersDefCli,TiersDefFou : String[17] ;
     RegimeDefaut       : String[3] ;
     DefautDDivers : String[17] ;
     DefautCDivers : String[17] ;
     DefautDivers : String[17] ;
     CptLettrDebit,CptLettrCredit : String[17] ;
     FBil : TFBil ;
     FCha : TFCha ;
     FPro : TFPro ;
     FExt : TFExt ;
     OnCumEdt,AlertePartiel,AlerteRegul,BouclerSaisieCreat : Boolean ;
     NumPlanRef,NbGC : Integer ;
     LgTableLibre  : TLgTableLibre ;
     NivTableLibre : TNivTableLibre ;
     OkTableLibre : TParamTableLibre ;
     BourreLibre : Char ;
     ImportRL,FromPCL,RecupPCL {récup PCL} : Boolean ;
     Mugler : Boolean ;
     SousPlanAxe : TSousPlanAxe ;
     CtrlRib     : Boolean ;
     GereSousPlan : Boolean ;
     SaisieTranche : Array[1..MaxAxe] of Boolean ;
     CatBud : TCatBud ;
     MultiSouche : Boolean ;
     // Sérialisation
     SerProdAna,SerProdCompta,SerProdBudget,SerProdImmo,SerProdEtebac,SerProdIcc,SerProdRevision, SerProdCCMP, SerProdCre, SerProdImmoPGI, SerProdGed: String ;
     SerProdTreso : string; {JP 27/01/04 : pour la seria de la trésorerie}
     SerProdCPPackAvance : String ; // CCS5 MODIF PACK AVANCE
     SerProdMultilingue : String ;  // module multi-lingue
     SerProdPackIFRS    : String ;  // Modif IFRS
     SerProdMess        : string;
     SerProdAgricole    : string;
     SerProdSuiviReg    : string;
     SerProdGCD         : string;
     SerProdDRF         : string;
     SerProdRIC         : string;
     SerProdDAS2        : string;
     SerProdExpertETEBAC: string;
     SerProdSCAN        : string;
     OkModAna,OkModCompta,OkModBudget,OkModImmo,OkModEtebac,OkModIcc : boolean;
     OkModRevision, OkModCCMP, OkModCre, OkModImmoPGI, OkModGed, OkModAgricole : boolean ;
     OkModTreso : Boolean; {JP 27/01/04 : pour la seria de la trésorerie}
     OkModCPPackAvance : Boolean ; // CCS5 MODIF PACK AVANCE
     OkModMultilingue  : Boolean ; // module multi-lingue
     OkModCPPackIFRS   : Boolean ;  // Modif IFRS
     OkModMessagerie   : boolean ;  // Messagerie (depuis le bureau)
     OkModSuiviReg     : boolean;
     OkModGCD          : boolean; // Gestion des Créances Douteuses
     OkModDRF          : boolean; // Détermination du résultat fiscal
     OkModRIC          : boolean; // Révision intégrée
     OkModDAS2         : boolean; // DAS2
     OkModExpertETEBAC : boolean; // ETEBAC Expert
     OkModSCAN         : boolean; // SCAN Expert
     // Immobilisations
     Exercices : Array [1..20] of TExoDate ;
     CpteCBInf,CpteCBSup : string[17];
     CpteLocInf,CpteLocSup : string[17];
     CpteDepotInf,CpteDepotSup : string[17];
     CpteImmoInf,CpteImmoSup : string[17];
     CpteFinInf,CpteFinSup : string[17];
     CpteAmortInf,CpteAmortSup : string[17];
     CpteDotInf,CpteDotSup : string[17];
     CpteExploitInf,CpteExploitSup : string[17];
     CpteDotExcInf,CpteDotExcSup : string[17];
     CpteRepExcInf,CpteRepExcSup : string[17];
     CpteVaCedeeInf,CpteVaCedeeSup : string[17];
     CpteDerogInf,CpteDerogSup : string[17];
     CpteProvDerInf,CpteProvDerSup : string[17];
     CpteRepDerInf,CpteRepDerSup : string[17];
     Specif : Array[0..9] Of Boolean ;
     ChargeOBImmo : boolean;
     //SpeedCum : Boolean ;
(*======= PFUGIER *)
     //SpeedObj:   Pointer ;
     //bSpeedFree: Boolean ;
     bByGenEtat: Boolean ;
(*======= PFUGIER *)
(* GP Pour menu POP interactif *)
     MPPop(*,SauveMpPop*) : tMPPop ;
     MPModifFaite : Boolean ;
(* GP Pour menu POP interactif *)
     ProfilUserC : TProfil ;
     ProfilCCMPExiste : Array[prCLient..prEtablissement] Of Boolean ;
// CA - 30/04/2002  {IFDEF CCMP}
     AccesMP : String ;
     SauveSA_SaisieEff : tSauveSA_SaisieEff ;
// CA - 30/04/2002 {ENDIF}
     RecupCegid : Boolean ;
     VNUMSOC : Integer ; // Pour l'exe EURO
     RecupLTL : Boolean ; // Pour récup le tout lyon
     RecupSISCOPGI : Boolean ; // Comme son nom l'indique
     TOBJalEffet : TOB ;
     RecupComSX : Boolean ;
     OldTeletrans : Boolean ;
     Cyber : Boolean ; // Pour récup Cybertech
     EnSerie : Boolean; // Pour traitement en série sans affichage à l' écran
     STOPRSP : Boolean ; // Comme son nom l'indique ! R(ecup) S(isco) P(GI)
//     FinEuro : Boolean ;
     TobCum : TOB ;
     UseTC : Boolean ;
     bAnalytique : Boolean; // Pour compte de résultat et SIG analytique
     IsBaseConso : Boolean ; // Pour identifier une base conso
     PointageJal : Boolean ;
// GC 26/069/2002
     StopEdition : Boolean ;
// FIN GC
     CCMP : tCCMP ;
//RR IFDEF CEGID
     CPIFDEFCEGID : boolean ;
//FIN RR
//SG6 21/01/05 Pour savoir dans quel mode analytique on se trouve
     AnaCroisaxe : boolean;
     TresoNoDossier : string; {JP 02/10/06 : NoDossier de la base Tréso}
   {$IFNDEF EAGLSERVER}
     BStatutLiasse  : TToolBarButton97; // GCO - 26/10/2007
     BStatutRevision : TToolBarButton97;
   {$ENDIF}
     RoleCompta : TCPRoleCompta;

     // GCO - 15/06/2007 - Record pour la Révision Intégrée Compta
      Revision : TCPREvision;

     Lequel : string ;
     LiaisonY2ViaShare : Boolean;
     END ;

{$IFNDEF NOVH}
{$IFDEF EAGLSERVER}
Type PLaVariableHalley = ^LaVariableHalley;
Type TLaVariableHalley = class
public
  FVH: PLaVariableHalley;
  constructor Create;
  destructor Destroy; override;
end;
{$ENDIF EAGLSERVER}
{$ENDIF !NOVH}

Type T_Eche = Record
              ModePaie         : String3 ;
              DateEche         : TDateTime ;
              MontantP         : Double ; // Pivot
              MontantD         : Double ; // Devise
//              MontantE         : Double ; // Euro
              Pourc            : Double ;
              ReadOnly         : Boolean ;
              Couverture       : Double ;
              CouvertureDev    : Double ;
//              CouvertureEuro   : Double ;
              CodeLettre       : String4 ;
              LettrageDev      : String1 ;
              DatePaquetMax    : TDateTime ;
              DatePaquetMin    : TDateTime ;
              DateRelance      : TDateTime ;
              NiveauRelance    : Integer ;
              EtatLettrage     : String3 ;
              DateValeur       : TDateTime ;
              {#TAVENC}
              TAV              : Array[1..5] of Double ;
              CodeTva          : String;
              Sens             : String;
              // ---
              PourQui          : string; // paiement direct sous traitant
              PourQuiLib       : string; // paiement direct sous traitant
{$IFDEF CHR}
              Folio            :integer;
              Mtencais         : Double;
              Datecreation     : TDatetime;
              Utilisateur      : string3;
{$ENDIF}

              end ;

    T_TabEche = Array[1..MaxEche] Of T_Eche ;

    T_ModeRegl = Record
                 Action       : TActionFiche ;
                 ModeInitial  : String3 ;
                 ModeFinal    : String3 ;
                 NbEche       : Integer ;
                 APartir      : String3 ;
                 ArrondirAu   : String3 ;
                 SeparePar    : String3 ;
                 PlusJour     : Integer ;
                 TotalAPayerP : Double ;
                 TotalAPayerD : Double ;
//                 TotalAPayerE : Double ;
                 JourPaiement1 : Integer ;
                 JourPaiement2 : Integer ;
                 CodeDevise   : String3 ;
                 Symbole      : String3 ;
                 TauxDevise   : Double ;
                 Quotite      : Double ;
                 Decimale     : Integer ;
                 DateFact     : TDateTime ;
                 DateFactExt  : TDateTime ;
                 DateBL       : TDateTime ;
                 Aux          : String17 ;
{$IFDEF GCGC}
                 EAN          : String17 ;
                 Bloque       : boolean;
{$ENDIF GCGC}
                 TauxEscompte : Double ;
                 ModeGuide    : boolean ;
                 EcartJours   : String ;
                 ModifTva     : boolean ;
                 TabEche      : T_TabEche ;
{$IFDEF CHR}
                 TabFolioP    : Array[1..5] of Double ;
                 TabFolioD    : Array[1..5] of Double ;
                 TabFolioE    : Array[1..5] of Double ;
{$ENDIF}
                 End ;

type T_Relance = Class
                 TypeRel,Famille : String3 ;
                 Groupe,NonEchu,Scoring : boolean ;
                 Delais  : Array[1..7] of integer ;
                 Modeles : Array[1..7] of String3 ;
                 End ;


Type tQuelProduit = (execcs3,exeCCS5, execcs7, exeCCIMPEXP, exeCCMP, exeCCAUTO,
                     exePGIMAJVER, ExeAutre,ExeCTS3, ExeCTS5, exeCIIMPEXP, exeCCSTD,
                     exeCOMSX, exeCCADM) ;




{$IFNDEF EAGLCLIENT}
Var ProcZoomGene    : TProcZoom ;
    ProcZoomTiers   : TProcZoom ;
    ProcZoomSection : TProcZoom ;
    ProcZoomJal     : TProcZoom ;
    ProcZoomBudGen  : TProcZoom ;
    ProcZoomBudSec  : TProcZoom ;
    ProcZoomBudJal  : TProcZoom ;
    ProcZoomCorresp : TProcZoom ;
    ProcZoomNatCpt  : TProcZoom ;
    ProcZoomRub     : TProcZoom ;
{$IFDEF GCGC}
    ProcZoomArticle : TProcZoom ;
{$ENDIF}
{$ENDIF}


{$IFNDEF NOVH}
{$IFNDEF EAGLSERVER}
Var VH : ^LaVariableHalley;
{$ELSE  !EAGLSERVER}
function VH: PLaVariableHalley;
{$ENDIF !EAGLSERVER}
{$ENDIF !NOVH}
{$IFNDEF EAGLSERVER}
procedure InitPopup( F : TForm) ;
procedure PurgePopup( PP : TPopupMenu ) ;
procedure PopZoom ( BM : TBitBtn ; POPZ : TPopupMenu ) ;
procedure PopZoom97 ( BM : TToolbarButton97 ; POPZ : TPopupMenu ) ;
function  CreerLigPop ( B : TBitBtn ; Owner : TComponent ; ShortC : Boolean ; C : Char ) : TMenuItem ;
{$ENDIF}
procedure NOMBREPEREXO (Exo : TExoDate ; Var PremMois,PremAnnee,NombreMois : Word) ;
Function  BourreLaDonc ( St : String ; LeType : TFichierBase ) : string ;
Function  BourreEtLess ( St : String ; LeType : TFichierBase ; ForceLg : Integer = 0) : string ;
function BourreOuTronque(St : String ; fb : TFichierBase) : String ;
Function  BourreLaDoncSurLaTable ( Code, St : String; ForceLg : Integer = 0 ) : string ;
function  BourreLaDoncSurLesComptes(st : string ; c : string = '') : string;
function  CompteDansLeIntervalle(Quoi, Deb, Fin : string) : boolean;
// Fonctions de conversions
Function  AxeToTz ( Ax : String ) : TZoomTable ;
function  AxeToDataType(CodeAxe : string) : string; {<=> AxeToTz pour les DataType et non plus les ZoomTable}
Function  NatureToTz (Nat : String) : TZoomTable ;
function  NatureToDataType(CodeTable : string) : string;{<=> AxeToTz pour les DataType et non plus les ZoomTable}
Function  StringToTz(s : String) : TZoomTable ;
Function  tzToNature( tz : TZoomTable ) : String ;
Function  tzToChampNature( tz : TZoomTable ; AvecPrefixe : Boolean) : String ;
Function  AxeToFb ( Axe : String ) : TFichierBase ;
Function  AxeToFbBud ( Axe : String ) : TFichierBase ;
Function  FbToAxe ( fb : TFichierBase ) : String ;
function StrToTA(psz : String) : TActionFiche;
function TAToStr(TA : TActionFiche) : String;

Function  BourreLess ( St : String ; LeType : TFichierBase ) : string ;
Function  FbSousPlantoFb ( fb : TFichierBase ) : TFichierBase ;
Procedure AvertirMultiTable (FTypeTable : String ) ;
Function  AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
function  fileTemp(const aExt: String): String;
Procedure ChangeMask (C : THNumEdit ; Decim : Integer ; Symbole : String) ;
Procedure ATTRIBSOCIETE( T : TDataSet ; P : String) ;
{$IFNDEF EAGLSERVER}
Function  InitPath : String ;
{$ENDIF}
procedure AfficheLeSolde (T : THNumEdit ; TD,TC : Double) ;
Procedure RempliExoDate (Var Exo : TExoDate ) ;
function  INSERE( st,st1 : string ; deb,long : integer) : string ;
function  BISSEXE (Annee : Word ) : Boolean ;
{$IFNDEF NOVH}
PROCEDURE CHARGEPARAMTABLELIBRE(St : String) ;
FUNCTION  CHARGESOCIETEHALLEY : Boolean ;
{$ENDIF}
{$IFNDEF NOVH}
Procedure ChargeStructureUnique ;
FUNCTION  CHARGEMAGHALLEY : Boolean ;
PROCEDURE CHARGETVATPF ;
procedure ChargeDossierTreso; {JP 02/10/06 : chargement du NoDossier de la base Tréso}
FUNCTION  TVA2ENCAIS(ModeTVA,Tva : String3 ; Achat : Boolean; RG : boolean=false; FarFae : boolean=false) : String ;
FUNCTION  TVA2CPTE(ModeTVA,Tva : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
FUNCTION  TVA2TAUX(ModeTVA,Tva : String3 ; Achat : Boolean) : Real ;
FUNCTION  TPF2ENCAIS(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
FUNCTION  TPF2CPTE(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
FUNCTION  TPF2TAUX(ModeTVA,Tpf : String3 ; Achat : Boolean) : Real ;
{$ENDIF}
FUNCTION  TOTDIFFERENT(X1,X2 : Double) : BOOLEAN ;
{$IFNDEF NOVH}
FUNCTION  HT2TVA ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  HT2TPF ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  HT2TTC ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  TTC2HT ( TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  TTC2TPF ( TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
FUNCTION  TTC2TVA (TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
{$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE}
Procedure ChgMaskChamp (C : TFloatField ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) ;
{$ENDIF}
Function  MontantToStr (Montant : Double ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) : string ;
Procedure CorrespToCombo(Var CB : THValComboBox ; Fb : TFichierBase );
Procedure CorrespToCodes(Plan : THValComboBox ; Var C1, C2 : TComboBox);
Procedure RuptureToCodes(Plan : THValComboBox ; Var C1, C2 : TComboBox ; Fb : TFichierBase );
Procedure CodesRuptSec(Var LCodes : TStringlist ; Plan,Axe : String) ;
Function  PrintSolde(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean) : String ;
Function  PrintSoldeFormate(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean ; FormMont : String) : String ;
Function  PrintSolde2(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean ; FormMont : THValCombobox) : String ;
Function  PrintEcart(TD,TC : DOUBLE ; Decim : Integer ; DebitPos : Boolean) : String ;
function  MoinsSymbole(LeSolde : String) : double ;
Function  SqlCptInterdit(LeChamp : String ; Var PourSql : String ; ZoneInterdit : TEdit) : Boolean ;
Function  Fourchette( Var St : String) : String ;
{$IFNDEF NOVH}
procedure ExoToEdDates(Exo : String3; ED1, ED2 : TControl); {Gestion du cas où Exo = '' sinon appel de ExoToDates}
procedure ExoToDates ( Exo : String3 ; ED1,ED2 : TControl ) ;
{$ENDIF} 
Procedure PremierDernier(fb : TFichierBase ; Var Cpt1,Cpt2 : String ) ;
Function  SQLPremierDernier(fb : TFichierbase ; Prem : Boolean) : String ;
Procedure PremierDernierRub(ZoomTable : TZoomTable ; SynPlus : String ; Var Cpt1,Cpt2 : String ) ;
Function  SQLPremierDernierRub(ZoomTable : TZoomTable ; SynPlus : String ; Prem : Boolean) : String ;
function  EstConfidentiel ( Conf : String ) : boolean ;
function  EstSQLConfidentiel ( StTable : String ; Cpte : String17 ) : boolean ;
{$IFNDEF NOVH}
Function  JaiLeDroitNature ( Fichier,Niveau,Action : integer ; Parle : boolean ) : boolean ;
Function  LWhereV8 : String ;
Procedure ListePeriode(Exo : String3 ; LLItems,LLValues : HTStrings ; Deb : Boolean; bAvecCloture : Boolean=TRUE) ;
{$ENDIF}
function  DateBudget(CPER : THValComboBox) : String ;
{$IFNDEF EAGLSERVER}
Function  SaisieLancer : Boolean ;
Function  SaisieFolioLancee : Boolean ;
{$ENDIF}
{$IFNDEF NOVH}
Function  VerifCorrespondance(Qui : Integer ; C1,C2 : String ) : Byte ;
{$ENDIF}
{$IFNDEF EAGLSERVER}
procedure ChargeSatellites(Pop : TPopupMenu ; Proc : TNotifyEvent) ;
{$ENDIF}
Function  CoherencePaysRegion (Pref : String ; T : TDataSet) : Boolean ;
{$IFNDEF NOVH}
PROCEDURE CHARGETABLELIBRE ;
{$ENDIF}
Function  QuelDateDeExo(CodeExo : String3 ; Var Exo : TExoDate) : Boolean ;
Procedure ChargeComboTableLibre(Cod : String ; DesValues,DesItems : HTStrings) ;
Procedure GetLibelleTableLibre(Pref : String ; LesLib : HTStringList) ;
Function  GetLibelleTabLibCommun(Cod : String) : String ;
{$IFNDEF NOVH}
PROCEDURE CHARGESOUSPLANAXE ;
PROCEDURE CHARGECATBUD ;
Procedure ChargeTablo(Laxe:String ; Var DT1,LT2 : TabByte ; Var CT3 : TabSt3 ; Var LiT4 : TabSt35) ;
function  RechargeParamSoc : boolean ;
function  CHARGEMAGEXO ( Parle : boolean ) : boolean ;
Function  GetVH(s : string) : string ;
{$ENDIF}
Function  QuelleCatBud(CodeCat : String) : TUneCatBud ;
Function  SousPlanCat(CodeCat : String ; SurJal : Boolean) : TSousPlanCat ;
Function  EuroOk : Boolean ;
Function  FinFrancs(DD : tDateTime) : Boolean ;
Procedure LibellesTableLibre ( PZ : TTabSheet ; NamL,NamH,Pref : String ) ;
{$IFNDEF EAGLSERVER}
Procedure LookLesDocks ( FF : TForm ) ;
{$ENDIF EAGLSERVER}
{$IFNDEF EAGLSERVER}
{$IFNDEF NOVH}
procedure UpdateSeries ;
{$ENDIF}
{$ENDIF EAGLSERVER}
procedure UpdateCombosLookUp ;
Function  AnalyseCompte(St : String ; fb : TFichierBase ; Exclu,OnTableLibre : Boolean ; RenvoieVide : Boolean = TRUE; DoDelete : Boolean = True; LesCptes :TStringList = nil) : String ; // 12102
{$IFNDEF NOVH}
Procedure ChargeHalleyUser ;
Function  EXRF ( CodeExo : String ) : String ;
{$ENDIF}
FUNCTION  QUELEXODTBUD(DD : TDateTime) : String ;
Function  ParamEuroOk : Boolean ;
Function  GetPeriode ( LaDate : TDateTime ) : integer ;
Function  IncPeriode ( Periode : integer ) : integer;
Function  GetDateTimeFromPeriode ( Periode : integer ) : TDateTime;
procedure DirDefault(Sauve : TSaveDialog ; FileName : String) ;
Function  PieceSurFolio(Jal : String) : Boolean ;
Function  AuMoinsUneImmo : Boolean ;
Procedure SetTousCombo ( TH : THValComboBox ) ;
(*======= PFUGIER *)
procedure MajBeforeDispatch(Num: Integer=0) ;
(*======= PFUGIER *)
Function LienS1 : Boolean ;
Function LienS3 : Boolean ;
Function LienS1S3 : Boolean ;
{$IFNDEF SPEC302}
{$IFNDEF NOVH}
Procedure ChargeMPACC ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure PositionneDomaineUser ( LeDom : THValComboBox ; IgnoreForce : Boolean = False; IgnoreVisible : Boolean = False; IgnoreDisabled : boolean=false) ;
Procedure PositionneEtabUser ( LeEtab : TControl ; IgnoreForce : Boolean = False; IgnoreVisible : Boolean = False; IgnoreDisabled : boolean=false) ;
Procedure PositionneDepotUser (LeDepot : TControl; IgnoreForce : Boolean = False; IgnoreVisible : boolean=false; IgnoreDisabled : boolean=false);
Procedure PositionneResponsableUser (LeResp : TControl; IgnoreForce : Boolean = False; IgnoreVisible : Boolean = False; IgnoreDisabled : boolean=false) ;
//
procedure LibelleEtablissement(CodeEtab: string; var lib1, lib2: string);
procedure LibelleDomaine(CodeDom: string; var lib1, lib2: string);

{$IFNDEF EAGLSERVER}
Procedure AGLPositionneDomaineUser ( Parms : array of variant ; nb : integer ) ;
{$ENDIF EAGLSERVER}
Function WhereProfilUser(Q : TQuery ; pf : tProfilTraitement) : String ;
Procedure ChargeProfilUser ;
Function  EtabForce : String ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure CreerDeviseTenue(LaMonnaie : String) ;
{$ENDIF}
Function  CompareTL(St1,St2 : String) : Boolean ;
{$IFNDEF NOVH}
Function  CptDansProfil(Gen,Aux : String ; Ind : tProfilTraitement) : Boolean ;
Function  MonProfilOk(Qui : tProfilTraitement) : Boolean ;
{$ENDIF}
Function _ChampExiste(NomTable,NomChamp : String) : Boolean ;
{$IFDEF EURO}
Procedure EcritProcEuro(St : String) ;
{$ENDIF}
//Function OkSynchro : Boolean ;
FUNCTION GoINSERE( st,st1 : string ; deb,long : integer) : string ;
{$IFNDEF NOVH}
Function HalteAuFeu : Boolean ;
Procedure DemandeStop(St : String) ;
Function EstSpecif(St : String) : Boolean ;
{$ENDIF}
function EstMonnaieIN ( CodeD : String3 ) : boolean ;
Function StopEcart (DD : tDateTime ; Dev : String = '') : Boolean ;
//Function StopDevise (DD : tDateTime ; Dev : String ; ModeOppose : Boolean=FALSE) : Boolean ;
Function EstbaseConso : Boolean ;
Function EstImmoPGI : Boolean ;

function EstComptaTreso : Boolean;
{$IFNDEF NOVH}
Function EstComptaSansAna : Boolean ;
function EstComptaIFRS : Boolean;
{$ENDIF}
Function EstComptaPackAvance : Boolean ;
// Fonctions communes avec la trésorerie
function calcIBAN	(pays : string ; RIB : string) : string ;
function calcRIB (pays : string ; banque : string ; guichet : string ; compte : string; cle : string) : string ;
{ Retraitement suite suppression contre-valeur }
{$IFDEF COMPTA}
//procedure SupprimeChampsEcartDansLesListes;
// function SupprimeEcartConversionEuro : boolean;
procedure ExportS7ImportS5Confidentialite;
{$ENDIF}
function IsDossierCabinet (NumDossier : String) : Boolean;
procedure ForceCascadeForms;

Function VERIFNIF_ES ( NIF : String ) : String ;

{$IFNDEF NOVH}
Procedure InitLaVariableHalley ;
{$IFNDEF EAGLSERVER}
Procedure ReleaseLaVariableHalley ();
{$ELSE}
Procedure ReleaseLaVariableHalley (RelpVH:PLaVariableHalley);
{$ENDIF}
{$ENDIF !NOVH}
function c_CheckProductSeriaPGI(Product: LongWord): integer; stdcall; external 'PGISERIA.DLL';
function IsSerialise(product : LongWord) : integer ;
function CIsValideDate(aDate : string) : Boolean;

function CEstPointageEnConsultationSurDossier : boolean;

function CGenereSQLConfidentiel( vStPrefixe : string ) : string;
{$IFNDEF EAGLSERVER}
function _IsDossierNetExpert : Boolean;
function BrancheParamSocAVirer : string;
{$ENDIF}

Function EstUneLigneCpt(St : String) : Boolean ;
function OkChar( St : String ) : Boolean ;
function TL_TIERSCOMPL_Actif : boolean;
// ajout me  01/09/2005
procedure MajLettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
function EstJoker(St: string): boolean;
// Ajout GCO - 24/01/2006
{$IFDEF COMPTA}
function CPEstTvaActivee : Boolean;
function CPGetMontantTVA(vStCodeTva : string ; vDateDebut, vDateFin : TDateTime) : Double;
Procedure MAJHistoparam (NomTable, Data :string);
function CPEstComptaAgricole : Boolean;
{$ENDIF}
{$IFDEF COMSX}
Procedure MAJHistoparam (NomTable, Data :string);
{$ENDIF}
// Fin GCO
{$IFNDEF EAGLSERVER}
procedure CPEnregistreLOG( vStCodeLog : string ); // GCO - 18/09/2006
// BVE 29.08.07 : Gestion du journal des évènements
{$IFDEF CERTIFNF}
function CPEnregistreJalEvent( TypeEvent : String ; Libelle : String ; BlocNote : String ) : integer ; overload;
function CPEnregistreJalEvent( TypeEvent : String ; Libelle : String ; BlocNote : HTStringList ) : integer ; overload;
{$ENDIF}
function GetRoleComptaUtilisateur( vUtilisateur : string ) : TCPRoleCompta;
{$ENDIF}

{$IFNDEF NOVH}
function JaiLeRoleCompta (rc : TCPRoleCompta) : boolean;
{$ENDIF}
function CPPresenceEtafi : boolean;
Function _GetParamSocSecur(St : String  ; ValDefaut : Boolean ; AvecMess : Boolean = FALSE) : Boolean ;
Procedure RetoucheHVCPoursaisie(TC : tControl) ;
function Entite : integer;
function OkMultiEntite: Boolean;
function GestionSequence:boolean;
function CPIncrementerCompteur(TypeSouche,CodeSouche:string;DD:tDateTime;var MasqueNum:string17;Nombre:integer=1):integer;
function CodeExo(DD:tDateTime):string;

type

 TZHalleyUser = Class(Tpersistent)
  private
   FGrpMontantMin     : double ;
   FGrpMontantMax     : double ;
   FJalAutorises      : string ;
   FDossier           : string ;
  public

   constructor Create ( vDossier : String = '' ) ; virtual ;
   procedure   ChargeHalleyUser ;

   property GrpMontantMin    : double read FGrpMontantMin ;
   property GrpMontantMax    : double read FGrpMontantMax ;
   property JalAutorises     : string read FJalAutorises ;
  end ;



 TZInfoCpta = Class
  private
   FSaisieTranche  : array [1..MaxAxe] of boolean ;
   FCpta           : array [fbAxe1..fbAux] of TInfoCpta ;
   FLgTableLibre   : TLgTableLibre ;
   FSousPlanAxe    : TSousPlanAxe ;
   FDossier        : string ;
   fMPACC           : TStrings ;

   function        GetCpta ( Value : TFichierBase ) : TInfoCpta ;
   function        GetLgTableLibre ( i , j : Integer) : Integer ;
   function        GetSousPlanAxe ( fb : TFichierBase ; vNumPlan : integer ) : TSousPlan ;
  public

  constructor Create( vDossier : String = '' ) ;
  destructor  Destroy; override ;

  procedure   ChargeSousPlanAxe ;
  procedure   ChargeLgDossier ;
  function    local : boolean ;
  Procedure   ChargeMPACC ;

  property Cpta        [ Value : TFichierBase]                    : TInfoCpta read GetCpta ;
  property LgTableLibre[ i , j : Integer]                         : integer   read GetLgTableLibre ;
  property SousPlanAxe [vAxe : TFichierBase ; vNumPlan : integer] : TSousPlan read GetSousPlanAxe ;
  property MPACC : TStrings read fMPACC write fMPACC;
 end;

  TZCatBud = class
  private
    FCatBud : array[1..MaxCatBud] of TUneCatBud;

    function GetCatBud(Value : Integer) : TUneCatBud;
  public
    constructor Create(vDossier : string = '');
    property    CatBud[aCatBud : Integer] : TUneCatBud read GetCatBud;
  end;

Function GetTenuEuro : Boolean;
Function GetRecupPCL : Boolean;
Function GetFromPCL : Boolean;
Function GetRecupComSx : Boolean;
Function GetRecupSISCOPGI : Boolean;
Function GetRecupLTL : Boolean;
Function GetRecupCegid : Boolean;
Function GetCPIFDEFCEGID : Boolean;


function GetInfoCpta ( Value : TFichierBase ) : TInfoCpta ;
function GetLgTableLibre(I, J : integer) : integer;
function GetEnCours : TExoDate ;
function GetNMoins2 : TExoDate ;
function GetPrecedent: TExoDate ;
function GetExoClo:  TTabExo ;
function GetSuivant : TExoDate ;
function GetCPExoRef : TExoDate ;
function GetEntree : TExoDate ;
function GetBourreLibre : Char ;
function GetGrpMontantMin : double ;
function GetGrpMontantMax : double ;
function GetJalAutorises : string ;
Function GetMultisouche : Boolean;
Function GetModCPPackAvance : Boolean;
Function GetAnaCroisaxe : Boolean;
function GetSousPlanAxe(fb : TFichierBase ; vNumPlan : integer) : TSousPlan ;
function GetCatBud(aCatBud : Integer) : TUneCatBud;
function QuelPaysLocalisation : string;
function GetExercices(n : Integer) : TExoDate ;

procedure CPStatutDossier ( vStDossier : string = '' ) ;
function GetInfoAcceptation  : TStrings ;

{b FP 21/02/2006}
type
  TCompensation = class
  public
    class function IsCompensation: Boolean; {Retourne True si un plan de compansation est utilisé}
    class function IsPlanCorresp(NumPlan: Integer): Boolean;  {Permet de savoir si un plan de correspondance est utilisé pour la compensation}
    class function GetSQLCorrespCompte(NatureCpt, Compte: String): String; {Nature compte C,1: Client - F,2: Fournisseur}
    class function IsNatureClient(Nature: String): Boolean;
    class function IsNatureFournisseur(Nature: String): Boolean;
    class function GetNomPlanCorrespondance(NatureCpt, Compte: String): String; {Retourne le plan de correspondance associé au compte}
    class function GetNumPlanCorrespondance: Integer;
    class function GetChampPlan: String;
    class function GetJalCompensation: String;
  end;

var VersionRef : string;

{e FP 21/02/2006}
implementation

uses
{$IFDEF EAGLSERVER}
{$IFDEF NOVH}
  ULibCpContexte ,
{$ENDIF}
{$ENDIF}
{$IFNDEF HVCL}
  ULibExercice ,
{$ENDIF}
  {$IFDEF AMORTISSEMENT}
  imEnt,
  {$ELSE}
   {$IFDEF PGIIMMO}
    imEnt,
   {$ENDIF}
  {$ENDIF}
  {$IFNDEF HVCL}
    {$IFNDEF EAGLCLIENT}
      {$IFNDEF EAGLSERVER}
      MajStruc,
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}
   ParamSoc
   {$IFDEF COMPTA}
    {$IFNDEF IMP}
   ,OuvreExo
   {$ENDIF}
   {$ENDIF}
   ,UtilPGI
{$IFDEF CHR}
  ,EntGC
{$ENDIF CHR}
	,uBTPVerrouilleDossier
  ,BTPGetVersions
  ,UBTPSequenceCpta
  ,UApplication
  ,UConnectWSCEGID 
   ;

{$IFDEF EAGLCLIENT}
{$IFDEF AUCASOU}
function GetModePCL ( st : string ) : string;
var CegidIni : TIniFile;
begin
  { Solution temporaire pour détecter le mode PCL }

  //Modif FV : Version 5 - Version 7
  //CegidIni := TIniFile.Create('c:\pgi00\env\cegidenv.ini');
  CegidIni := TIniFile.Create(TCBPPath.GetCegidDistri & '\env\CegidEnv.Ini');

  if CegidIni.ReadString('General', 'ModePCL', '1') = '1'
  then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
  CegidIni.Free;
end;
{$ENDIF}
{$ENDIF}

function IsSerialise(product : LongWord) : integer ;
begin
  result := c_CheckProductSeriaPGI(Product);
end;

Function LienS1 : Boolean ;
BEGIN Result:=(GetParamSocSecur('SO_CPLIENGAMME','')='S1') ; END ;

Function LienS3 : Boolean ;
BEGIN Result:=(GetParamSocSecur('SO_CPLIENGAMME','')='S3') ; END ;

Function LienS1S3 : Boolean ;
BEGIN Result:=(LienS1) or (LienS3) ; END ;

{$IFNDEF NOVH}
Function  EXRF ( CodeExo : String ) : String ;
Var Ex : String ;
BEGIN
Result:=CodeExo ;
if Not (ctxPCL in V_PGI.PGIContexte) then Exit ;
Ex:=VH^.CPExoRef.Code ; if Ex='' then Exit ;
if ((Ex=VH^.Encours.Code) or (Ex=VH^.Suivant.Code)) then Result:=Ex ;
END ;
{$ENDIF}

Procedure SetTousCombo ( TH : THValComboBox ) ;
BEGIN
if TH=Nil then Exit ;
if Not TH.Vide then Exit ;
if TH.Values.Count<=0 then Exit ;
TH.ItemIndex:=0 ;
END ;

Procedure ChargeTablo(Laxe:String ; Var DT1,LT2 : TabByte ; Var CT3 : TabSt3 ; Var LiT4 : TabSt35) ;
Var Q : TQuery ;
    i : Byte ;
BEGIN
FillChar(DT1,SizeOF(DT1),#0) ; FillChar(LT2,SizeOF(LT2),#0) ;
FillChar(CT3,SizeOF(CT3),#0) ; FillChar(LiT4,SizeOF(LiT4),#0) ;
Q:=OpenSQL('SELECT SS_AXE,SS_SOUSSECTION,SS_DEBUT,SS_LONGUEUR,SS_LIBELLE FROM STRUCRSE WHERE SS_AXE="'+Laxe+'" Order by SS_DEBUT',TRUE,-1,'',true);
i:=1 ;
While Not Q.EOF do
  BEGIN
   CT3[i]:=Q.Fields[1].AsString ; DT1[i]:=Q.Fields[2].AsInteger ;
   LT2[i]:=Q.Fields[3].AsInteger ; LiT4[i]:=Q.Fields[4].AsString ;
   Q.Next ; Inc(i) ;
  END ;
Ferme(Q) ;
END ;

Function  GetLibelleTabLibCommun(Cod : String) : String ;
Var QLoc : TQuery ;
    lStDossier : string ;
BEGIN
// Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then lStDossier := TableToBase( 'NATCPTE' )
    else lStDossier := '' ;
QLoc:=OpenSelect('Select CC_LIBELLE From CHOIXCOD Where CC_TYPE="NAT" And CC_CODE ="'+Cod+'"',lStDossier) ;
Result:=QLoc.Fields[0].AsString ;
Ferme(QLoc) ;
END ;

Procedure GetLibelleTableLibre(Pref : String ; LesLib : HTStringList) ;
Var QLoc : TQuery ;
    lStDossier : string ;
BEGIN
// Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then lStDossier := TableToBase( 'NATCPTE' )
    else lStDossier := '' ;
QLoc:=OpenSelect('Select CC_LIBELLE,CC_ABREGE,CC_CODE From CHOIXCOD Where CC_TYPE="NAT" And CC_CODE Like"'+Pref+'%" Order by CC_CODE',lStDossier) ;
LesLib.Clear ;
While Not QLoc.Eof do
   BEGIN LesLib.Add(QLoc.Fields[0].AsString+';'+QLoc.Fields[1].AsString) ; QLoc.Next ; END ;
Ferme(QLoc) ;
END ;

Procedure ChargeComboTableLibre(Cod : String ; DesValues,DesItems : HTStrings) ;
Var QLoc : TQuery ;
    lStDossier : string ;
BEGIN
DesValues.Clear ; DesItems.Clear ;
if Cod='' then Exit ;
// Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then lStDossier := TableToBase( 'NATCPTE' )
    else lStDossier := '' ;
QLoc:=OpenSelect('Select CC_CODE,CC_LIBELLE From CHOIXCOD Where CC_TYPE="NAT" And CC_CODE Like "'+Cod+'%"',lStDossier, True, 'CHOIXCOD') ;
While Not QLoc.Eof do
   BEGIN
   DesValues.Add(QLoc.Fields[0].AsString) ; DesItems.Add(QLoc.Fields[1].AsString) ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

Function CoherencePaysRegion (Pref : String ; T : TDataSet) : Boolean ;
Var StP,StR : String ;
BEGIN
Result:=True ;
if T.FindField(Pref+'PAYS')=Nil then Exit ;
if T.FindField(Pref+'DIVTERRIT')=Nil then Exit ;
StP:=T.FindField(Pref+'PAYS').AsString ; StR:=T.FindField(Pref+'DIVTERRIT').AsString ;
if (StR='') and (StP='') then Exit ;
if (StR='') and (StP<>'') then Exit ;
if (StR<>'') and (StP='') then BEGIN Result:=False ; Exit ; END ;
Result:=PresenceComplexe('REGION',['RG_PAYS','RG_REGION'],['=','='],[StP,StR],['S','S']) ;
END ;

{$IFNDEF NOVH}
Function VerifCorrespondance(Qui : Integer ; C1,C2 : String ) : Byte ;
Var St : String ;
BEGIN
Result:=0  ;
Case Qui of
   1 : St:='GE' ; 2 : St:='AU' ; 3 : St:='A1' ; 4 : St:='A2' ;
   5 : St:='A3' ; 6 : St:='A4' ; 7 : St:='A5' ; 8 : St:='BU' ;
   Else Exit ;
   END ;
if ((VH^.VerifCorresp[Qui,1]) and (C1<>'')) then
   if Not PresenceComplexe ('CORRESP',['CR_TYPE','CR_CORRESP'],['=','='],[St+'1',C1],['S','S']) then Result:=1 ;
if ((VH^.VerifCorresp[Qui,2]) and (C2<>'') AND (Result=0)) then
   if Not PresenceComplexe ('CORRESP',['CR_TYPE','CR_CORRESP'],['=','='],[St+'2',C2],['S','S']) then Result:=2 ;
END ;
{$ENDIF}

// Regarde si la saisie ou la saisie tréso est lancée
{$IFNDEF EAGLSERVER}
Function SaisieLancer : Boolean ;
Var i : Integer ;
BEGIN
Result:=False ;
for i:=0 to Screen.FormCount-1 do
    if (UpperCase(Screen.Forms[i].Name)='FSAISIE') Or (UpperCase(Screen.Forms[i].Name)='FSAISIETR')
    Or (UpperCase(Screen.Forms[i].Name)='FSAISIEEFF') then BEGIN Result:=True ; Break ; END ;
END ;

Function SaisieFolioLancee : Boolean ;
Var i : Integer ;
BEGIN
Result:=False ;
for i:=0 to Screen.FormCount-1 do
    if UpperCase(Screen.Forms[i].Name)='FSAISIEFOLIO' then BEGIN Result:=True ; Break ; END ;
END ;


procedure ChargeSatellites(Pop : TPopupMenu ; Proc : TNotifyEvent) ;
var m1,m2      : TMenuItem ;
    SatFile    : TIniFile ;
    FSect      : TStringList ;
    i,j,k      : Integer ;
    NomSection : String ;
    ts         : TStrings ;
    s          : String ;
begin
PurgePopup(Pop) ;
ts := TStringList.Create ;
SatFile :=TIniFile.Create('HALSAT.INI');
FSect:=TStringList.Create ;
SatFile.ReadSections(FSect) ;
for i:=0 to FSect.Count-1 do
   begin
   NomSection := FSect[i] ;
   m1:= TMenuItem.Create(Pop) ;
   m1.Caption := NomSection ;
   Pop.Items.Add(m1) ;
   ts.Clear ;
   SatFile.ReadSectionValues(NomSection,ts) ;
   for j := 0 to ts.Count - 1 do
      begin
      s := ts[j] ;
      k := Pos('=',ts[j]) ;
      s := Copy(s,k+1,Length(s)-k) ;
      m2 := TMenuItem.Create(Pop) ;
      m2.Caption := ReadTokenSt(s) ;
      m2.Hint    := ReadTokenSt(s) ;
      m2.OnClick := Proc ;
      m1.Add(m2) ;
      end ;
   end ;
SatFile.Free ; ts.Free ; FSect.Free ;
end ;
{$ENDIF}

function  EstConfidentiel ( Conf : String ) : boolean ;
BEGIN
Result:=False ;
if V_PGI.Confidentiel='1' then Exit ;
if ((Conf='0') or (Conf='') or (Conf='-')) then Exit ;
Result:=True ;
END ;

function  EstSQLConfidentiel ( StTable : String ; Cpte : String17 ) : boolean ;
Var Q : TQuery ;
BEGIN
Result:=False ;
if V_PGI.Confidentiel='1' then Exit ;
if StTable='GENERAUX' then Q:=OpenSQL('Select G_CONFIDENTIEL from GENERAUX Where G_GENERAL="'+Cpte+'"',True,-1,'',true) else
if StTable='TIERS' then Q:=OpenSQL('Select T_CONFIDENTIEL from TIERS Where T_AUXILIAIRE="'+Cpte+'"',True,-1,'',true) else exit ;
if Not Q.EOF then Result:=EstConfidentiel(Q.Fields[0].AsString) else Result:=False ;
Ferme(Q) ;
END ;


{================================================================}
{$IFNDEF EAGLSERVER}
Function InitPath : String ;
BEGIN
Result:=ExtractFilePath(Application.exename) ;
if Length(Result)>3 then Delete(Result,Length(Result),1)
END ;
{$ENDIF}
{================================================================}
PROCEDURE ATTRIBSOCIETE( T : TDataSet ; P : String) ;
Var F : TField ;
BEGIN
F:=T.FindField(P+'_SOCIETE') ; if F=Nil then exit ;
F.AsString:=V_PGI.CodeSociete ;
END ;

function fileTemp(const aExt: String): String;
var
  Buffer: array[0..1023] of Char;
  aFile : String;
begin
  GetTempPath(Sizeof(Buffer)-1,Buffer);
  GetTempFileName(Buffer,'TMP',0,Buffer);
  SetString(aFile, Buffer, StrLen(Buffer));
  Result:=ChangeFileExt(aFile,aExt);
  RenameFile(aFile,Result);
end;

{=============================================================================}
FUNCTION BISSEXE (Annee : Word ) : Boolean ;
BEGIN
Result:=isLeapYear(Annee) ;
END ;

{============================================================================}
FUNCTION INSERE( st,st1 : string ; deb,long : integer) : string ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.ChargeMagExo') ;
{$ENDIF}
Insere:=copy(st,1,deb-1)+copy(st1,1,Long)+copy(st,deb+long,length(st)-long) ;
END ;

{============================================================================}
Procedure ChangeMask (C : THNumEdit ; Decim : Integer ; Symbole : String) ;
Var j : integer ;
BEGIN
C.Decimals:=Decim ; C.CurrencySymbol:=Trim(Symbole) ;
if Decim=0 then C.Masks.PositiveMask:='#,##0' else C.Masks.PositiveMask:='#,##0.' ;
for j:=1 to Decim do C.Masks.PositiveMask:=C.Masks.PositiveMask+'0' ;
C.Masks.ZeroMask:=C.Masks.PositiveMask ;
C.Masks.NegativeMask:='-'+C.Masks.PositiveMask ;
END ;

(*======================================================================*)
procedure AfficheLeSolde (T : THNumEdit ; TD,TC : Double) ;
BEGIN
if ((TD=TC) or (T.Tag=3)) then
   BEGIN
   T.NumericType:=ntGeneral ;
   if T.Value<>TD-TC then T.Value:=TD-TC ;
   END else if Abs(TD)>=Abs(TC) then
   BEGIN
   T.Debit:=True ; T.NumericType:=ntDC ;
   if T.Value<>TD-TC then T.Value:=TD-TC ;
   END else if Abs(TD)<Abs(TC) then
   BEGIN
   T.Debit:=False ; T.NumericType:=ntDC ;
   if T.Value<>TC-TD then T.Value:=TC-TD ;
   END ;
END ;

{============================================================================}
Procedure RempliExoDate (Var Exo : TExoDate ) ;
BEGIN
if Exo.Code=GetPrecedent.Code then
   BEGIN
   Exo.Deb:=GetPrecedent.Deb ; Exo.Fin:=GetPrecedent.Fin ;
   Exo.NombrePeriode:=GetPrecedent.NombrePeriode ;
   END else
   if Exo.Code=GetEnCours.Code then
      BEGIN
      Exo.Deb:=GetEnCours.Deb ; Exo.Fin:=GetEnCours.Fin ;
      Exo.NombrePeriode:=GetEnCours.NombrePeriode ;
      END else
      if Exo.Code=GetSuivant.Code then
         BEGIN
         Exo.Deb:=GetSuivant.Deb ; Exo.Fin:=GetSuivant.Fin ;
         Exo.NombrePeriode:=GetSuivant.NombrePeriode ;
         END
         else // ajout ME 02/10/2003 pour les exercice n-2
              QuelDateDeExo(Exo.Code,Exo) ;
END ;



{=====================================================================}
{$IFNDEF NOVH}
PROCEDURE CHARGETABLELIBRE ;
Var Q : TQuery ;
    CO,CL : String ;
    StCode,StLib : String ;
    Lg,i,j : Integer ;
    lStDossier : string ;
BEGIN
{Longueurs}
Fillchar(VH^.LgTableLibre,SizeOf(VH^.LgTableLibre),#0) ;

// Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then lStDossier := TableToBase( 'NATCPTE' )
    else lStDossier := '' ;

  Q:=OpenSelect('SELECT CC_CODE,CC_LIBRE,CC_TYPE FROM CHOIXCOD WHERE CC_TYPE="NAT" ORDER BY CC_TYPE,CC_CODE',lStDossier, True, 'CHOIXCOD') ;
  While Not Q.EOF Do
    BEGIN
    CO:=Q.Fields[0].AsString ; CL:=Q.Fields[1].AsString ;
    If OkChar(CL) Then
      BEGIN
      Lg:=StrToInt(CL) ;
      Case CO[1] Of
             'G' : i:=1 ; 'T' : i:=2 ; 'S' : i:=3 ; 'B' : i:=4 ;
             'D' : i:=5 ; 'E' : i:=6 ; 'A' : i:=7 ; 'U' : i:=8 ;
             'I' : i:=9 ;
             Else i:=0 ;
         END ;
       If i>0 Then BEGIN j:=StrToInt(Copy(CO,2,2))+1 ; VH^.LgTableLibre[i,j]:=Lg ; END ;
       END ;
    Q.Next ;
    END ;
  Ferme(Q) ;

{Longueurs}
{ Non pris en compte sur Immo }
for i:=1 to 5 do for j:=1 to 10 do VH^.NivTableLibre[i,j]:='XXX' ;
if True then
   BEGIN
   Q:=OpenSelect('SELECT CC_ABREGE, CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="NNC" AND CC_CODE="'+Inttostr(V_PGI.UserGrp)+'"',lStDossier, TRUE,'CHOIXCOD') ;
   While Not Q.EOF Do
     BEGIN
     StCode:=Q.Fields[0].AsString ;
     Case StCode[1] of
        'G' : i:=1 ; 'T' : i:=2 ; 'S' : i:=3 ; 'B' : i:=4 ;
        'D' : i:=5 ; 'E' : i:=6 ; 'A' : i:=7 ; 'U' : i:=8 ;
        Else i:=0 ;
      END ;
     if i<>0 then
        BEGIN
        StLib:=Q.Fields[1].AsString ; StLib:=Format_String(StLib,30) ;
        for j:=1 to 10 do VH^.NivTableLibre[i,j]:=Copy(StLib,1+3*(j-1),3) ;
        END ;
     Q.Next ;
     END ;
   Ferme(Q) ;
   END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
PROCEDURE CHARGEPARAMTABLELIBRE(St : String) ;
Var l,i : Integer ;
    St1 : String ;
BEGIN
Fillchar(VH^.OkTableLibre,SizeOf(VH^.OkTableLibre),#0) ;
If St='' Then Exit ;
if True then
   BEGIN
   l:=Pos('A',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,4) ;  For i:=1 To 4  Do If St1[i]='X' Then VH^.OkTableLibre[tlEcrAna,i]:=TRUE ; END ;
   l:=Pos('B',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,10) ; For i:=1 To 10 Do If St1[i]='X' Then VH^.OkTableLibre[tlCptBud,i]:=TRUE ; END ;
   l:=Pos('D',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,10) ; For i:=1 To 10 Do If St1[i]='X' Then VH^.OkTableLibre[tlSectBud,i]:=TRUE ; END ;
   l:=Pos('E',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,4) ;  For i:=1 To 4  Do If St1[i]='X' Then VH^.OkTableLibre[tlEcrGen,i]:=TRUE ; END ;
   l:=Pos('G',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,10) ; For i:=1 To 10 Do If St1[i]='X' Then VH^.OkTableLibre[tlGene,i]:=TRUE ; END ;
   l:=Pos('S',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,10) ; For i:=1 To 10 Do If St1[i]='X' Then VH^.OkTableLibre[tlSect,i]:=TRUE ; END ;
   l:=Pos('T',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,10) ; For i:=1 To 10 Do If St1[i]='X' Then VH^.OkTableLibre[tlAux,i]:=TRUE ; END ;
   l:=Pos('U',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,4) ;  For i:=1 To 4  Do If St1[i]='X' Then VH^.OkTableLibre[tlEcrBud,i]:=TRUE ; END ;
   l:=Pos('I',St) ; If l>0 Then BEGIN St1:=Copy(St,l+1,4) ;  For i:=1 To 4  Do If St1[i]='X' Then VH^.OkTableLibre[tlImmo,i]:=TRUE ; END ;
   END ;
END ;
{$ENDIF}

{=====================================================================}
{$IFNDEF SPEC302}
{$IFNDEF NOVH}
Procedure ChargeMPACC ;
Var Q : TQuery ;
    i : integer ;
    St : String ;
BEGIN
VH^.MPACC.Clear ;
Q:=OpenSQl('SELECT MP_MODEPAIE, MP_CODEACCEPT, MP_CATEGORIE, MP_LETTRECHEQUE, MP_LETTRETRAITE FROM MODEPAIE',True,-1,'',true);
While Not Q.EOF do
   BEGIN
   St:='' ; for i:=0 to 4 do St:=St+Q.Fields[i].AsString+';' ;
   VH^.MPACC.Add(St) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;
{$ENDIF}
Procedure TZInfoCpta.ChargeMPACC ;
Var Q : TQuery ;
    i : integer ;
    St : String ;
BEGIN
MPACC.Clear ;
Q:=OpenSQl('SELECT MP_MODEPAIE, MP_CODEACCEPT, MP_CATEGORIE, MP_LETTRECHEQUE, MP_LETTRETRAITE FROM MODEPAIE',True,-1,'',true);
While Not Q.EOF do
   BEGIN
   St:='' ; for i:=0 to 4 do St:=St+Q.Fields[i].AsString+';' ;
   MPACC.Add(St) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure ChargeParamsCpta ;
BEGIN
if VH^.PasParSoc then
   BEGIN
   VH^.JalVTP:='' ; VH^.JalATP:='' ; VH^.OuiTP:=False ; VH^.BouclerSaisieCreat:=True ;
   VH^.RegimeDefaut:='' ; VH^.TiersDefSal:='' ; VH^.TiersDefDiv:='' ;
   VH^.TiersDefCli:='' ; VH^.TiersDefFou:='' ;
   VH^.JalRepBalAN:='' ; VH^.AlertePartiel:=False ;
   VH^.EquilAnaODA:=False ; VH^.AlerteRegul:=False ;
   VH^.JalLookup:=False ; VH^.EtabLookup:=False ;
   VH^.AutoExtraSolde:=False ;
   VH^.CPRefLettrage:='REF' ; VH^.CPRefPointage:='REF' ;
   END else
   BEGIN
{$IFNDEF SPEC302}
   if GetParamSocSecur('SO_IFDEFCEGID',False) then
       VH^.CPIFDEFCEGID := True;

   VH^.JalVTP:=GetParamSocSecur('SO_JALVTP','') ; VH^.JalATP:=GetParamSocSecur('SO_JALATP','') ;
   VH^.OuiTP:=GetParamSocSecur('SO_OUITP','') ;
   //SG6 14/01/05 FQ 15229
   //VH^.EquilAnaODA:=GetParamSocSecur('SO_EQUILANAODA') ;
   VH^.JalLookup:=GetParamSocSecur('SO_JALLOOKUP','') ; VH^.EtabLookup:=GetParamSocSecur('SO_ETABLOOKUP','') ;
   VH^.CPRefLettrage:=GetParamSocSecur('SO_CPREFLETTRAGE','') ;
   //SG6 14/01/05 FQ 15229
   VH^.EquilAnaODA:=GetParamSocSecur('SO_EQUILANAODA','') ;
   VH^.AlertePartiel:=GetParamSocSecur('SO_ALERTEPARTIEL',False) ;
   VH^.RegimeDefaut:=GetParamSocSecur('SO_REGIMEDEFAUT','') ;
   VH^.TiersDefSal:=GetParamSocSecur('SO_SALATTEND','') ;
   VH^.TiersDefDiv:=GetParamSocSecur('SO_DIVATTEND','') ;
   VH^.TiersDefCli:=GetParamSocSecur('SO_CLIATTEND','') ;
   VH^.TiersDefFou:=GetParamSocSecur('SO_FOUATTEND','') ;
   VH^.JalRepBalAN:=GetParamSocSecur('SO_JALREPBALAN','') ;
   VH^.AlerteRegul:=GetParamSocSecur('SO_ALERTEREGUL',False) ;
   VH^.BouclerSaisieCreat:=GetParamSocSecur('SO_BOUCLERSAISIECREAT',False) ;
   VH^.AutoExtraSolde:=GetParamSocSecur('SO_AUTOEXTRASOLDE',False) ;
 {$IFNDEF SPEC350}
 {$IFNDEF EURO}
   VH^.CPLastSaisie:=GetParamSocSecur('SO_CPLASTSAISIE', iDate1900) ;
   VH^.CPExoRef.Code:=GetParamSocSecur('SO_CPEXOREF','') ;
 {$ENDIF}
 {$ENDIF}
{$ENDIF}
   END ;
{$IFNDEF SPEC302}
ChargeMPACC ;
{$ENDIF}
END ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure ChargeParamsImmo ;
BEGIN
// CA - 29/11/2004 - on charge systématiquement le contexte immo, ça évitera des problèmes ...
  //  if V_PGI.VersionDemo then Exit ;
  {$IFDEF AMORTISSEMENT}
  ChargeVHImmo;
  {$ELSE}
   {$IFDEF PGIIMMO}
   ChargeVHImmo;
   {$ENDIF}
  {$ENDIF}
  if VH^.PasParSoc then Exit ;
END ;
{$ENDIF}

Function QuelPaysLocalisation : String ;
Begin
  Result:=GetParamSocSecur('SO_PAYSLOCALISATION','') ;
  { FQ 19484
  if (Result<>CodeISOFR) and (Result<>CodeISOES) then
     Result:=CodeISOFR ;}
  if Result = '' then
     Result := CodeISOFR;
  {END FQ 19484 }
End ;

{$IFNDEF NOVH}
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 31/01/2007
Modifié le ... : 31/01/2007
Description .. : Charge le rôle comptable pour l'utilisateur courant
Mots clefs ... :
*****************************************************************}
procedure ChargeRoleMetier;
var
  Q : TQuery;
  ii, rc : integer;
begin
  rc := -1;
  // Chargement des menus : exclusion de la 'mère' avec mn_3 <> 0
  Q := OpenSQL('Select MN_TAG, MN_LIBELLE, MN_ACCESGRP FROM MENU WHERE MN_1=360 AND MN_3<>0 ORDER BY MN_TAG',
    True, -1, 'MENU', True);
  while not Q.EOF do
  begin
    ii := TrouveAcces(Q.FindField('MN_ACCESGRP').AsString, Q.FindField('MN_LIBELLE').AsString);
    if (ii=2) then break else Inc (rc);
    Q.Next;
  end;
  Ferme(Q);
  case rc of
    -1,0 : VH^.RoleCompta := rcCollaborateur;
    1 : VH^.RoleCompta := rcReviseur;
    2 : VH^.RoleCompta := rcSuperviseur;
    3 : VH^.RoleCompta := rcExpert;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFNDEF EAGLSERVER}
function GetRoleComptaUtilisateur( vUtilisateur : string ) : TCPRoleCompta;
var Q : TQuery;
    ii, rc : integer;
    lOldUserGrp      : integer;
    lOldGrpsDelegues : string;
    lStGrp : string;
begin
  rc := -1;

  lOldUserGrp      := V_Pgi.UserGRP;
  lOldGrpsDelegues := V_Pgi.GrpsDelegues;

  V_Pgi.GrpsDelegues := GetColonneSQL('UTILISAT', 'US_GRPSDELEGUES', 'US_UTILISATEUR = "' + vUtilisateur + '"');

  lStGrp := GetColonneSQL('UTILISAT', 'US_GROUPE', 'US_UTILISATEUR = "' + vUtilisateur + '"');
  V_Pgi.UserGrp := StrToInt(GetColonneSQL('USERGRP', 'UG_NUMERO', 'UG_GROUPE = "' + lStGrp + '"'));

  // Chargement des menus : exclusion de la 'mère' avec mn_3 <> 0
  Q := OpenSQL('SELECT MN_TAG, MN_LIBELLE, MN_ACCESGRP FROM MENU ' +
               'WHERE MN_1 = 360 AND MN_3 <> 0 ORDER BY MN_TAG', True, -1, 'MENU', True);

  while not Q.EOF do
  begin
    ii := TrouveAcces(Q.FindField('MN_ACCESGRP').AsString, Q.FindField('MN_LIBELLE').AsString);
    if (ii=2) then break else Inc (rc);
    Q.Next;
  end;
  Ferme(Q);

  case rc of
    -1,0 : Result := rcCollaborateur;
    1    : Result := rcReviseur;
    2    : Result := rcSuperviseur;
    3    : Result := rcExpert;
    else Result := rcCollaborateur;
  end;

  V_Pgi.UserGRP      := lOldUserGrp;
  V_Pgi.GrpsDelegues := lOldGrpsDelegues;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 31/01/2007
Modifié le ... :   /  /
Description .. : Teste le rôle en comptabilité pour l'utilisateur courant.
Suite ........ : La fonction renvoie VRAI si l'utilisateur courant a le rôle
Suite ........ : interrogé en paramètre
Mots clefs ... :
*****************************************************************}
function JaiLeRoleCompta (rc : TCPRoleCompta) : boolean;
begin
  Result := (rc <= VH^.RoleCompta);
end;
{$ENDIF}

{$IFDEF CCMP}
Procedure ChargeParamsCCMP ;
BEGIN
VH^.CCMP.LotCli:=GetParamSocSecur('SO_LOTCLICCMP',False) ;
VH^.CCMP.LotFou:=GetParamSocSecur('SO_LOTFOUCCMP',False) ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure ChargeJalEffet ;
Var Q : TQuery ;
    St : String ;
BEGIN
if VH^.TOBJalEffet<>Nil then BEGIN VH^.TOBJalEffet.Free ; VH^.TOBJalEffet:=TOB.Create('',Nil,-1) ; END ;
St:='SELECT J_JOURNAL, J_CONTREPARTIE, J_NATUREJAL, G_NATUREGENE, G_COLLECTIF, G_POINTABLE, G_LETTRABLE, '+
    'G_VENTILABLE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 '+
    ' FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL WHERE J_EFFET="X" AND J_CONTREPARTIE<>""' ;
Q:=OpenSQL(St,TRUE, -1,'JOURNAL',true);
VH^.TOBJalEffet.LoadDetailDB('ECRITURE','','',Q,FALSE,FALSE) ;
Ferme(Q) ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Function ChargeInfosSociete : Boolean ;
Var Stb,St: String ;
BEGIN
Result:=False ;
if GetParamSocSecur('SO_SOCIETE','')<>'' then V_PGI.CodeSociete:=GetParamSocSecur('SO_SOCIETE','') ;
V_PGI.NomSociete:=GetParamSocSecur('SO_LIBELLE','') ;
//  BBI Web services
//{$IFNDEF EAGLSERVER}
V_PGI.DevisePivot:=GetParamSocSecur('SO_DEVISEPRINC','') ;
V_PGI.SymbolePivot:='' ;
V_PGI.DeviseFongible:=V_PGI.DevisePivot ;
VH^.SymboleFongible:='' ;
//{$ENDIF}
//  BBI Fin Web services

{$IFNDEF EAGLSERVER}
V_PGI.LogSOC:=GetParamSocSecur('SO_SUIVILOG',False) ;
{$ENDIF}
VH^.NbjEcrAV:=GetParamSocSecur('SO_NBJECRAVANT',0) ; VH^.NbjEcrAP:=GetParamSocSecur('SO_NBJECRAPRES',0) ;
VH^.NbjEchAV:=GetParamSocSecur('SO_NBJECHAVANT',0) ; VH^.NbjEchAP:=GetParamSocSecur('SO_NBJECHAPRES',0) ;
VH^.CtrlRib:=GetParamSocSecur('SO_VERIFRIB',False) ;
//VH^.RegleEquilSais:=GetParamSocSecur('SO_REGLEEQUILSAIS') ;
{$IFNDEF EAGLSERVER}
V_PGI.Furtivite:=GetParamSocSecur('SO_FURTIVITE',0) ;

{ AJOUT CA 24/10/2005 - Correction des problèmes de synchronisation des
  OkDecV, OkDecP et OkDecQ entre les tables SOCIETE et PARAMSOC }
if V_PGI.OkDecV <> GetParamSocSecur('SO_DECVALEUR',0) then
  ExecuteSQL ('UPDATE SOCIETE SET SO_DECVALEUR='+IntToStr(GetParamSocSecur('SO_DECVALEUR',0)));
if V_PGI.OkDecQ <> GetParamSocSecur('SO_DECQTE',0) then
  ExecuteSQL ('UPDATE SOCIETE SET SO_DECQTE='+IntToStr(GetParamSocSecur('SO_DECQTE',0)));
if V_PGI.OkDecP <> GetParamSocSecur('SO_DECPRIX',0) then
  ExecuteSQL ('UPDATE SOCIETE SET SO_DECPRIX='+IntToStr(GetParamSocSecur('SO_DECPRIX',0)));

V_PGI.OkDecV:=GetParamSocSecur('SO_DECVALEUR',0) ;
V_PGI.OkDecQ:=GetParamSocSecur('SO_DECQTE',0) ;
V_PGI.OkDecP:=GetParamSocSecur('SO_DECPRIX',0) ;
{$ENDIF}
{Euro}
V_PGI.OkDecE:=2 ;
{$IFNDEF EAGLSERVER}
(* // SBO 01/07/2007 : FQ20888 --> Il ne faut plus charger le taux de parité des devises fongibles !
V_PGI.TauxEuro:=GetParamSocSecur('SO_TAUXEURO',0) ;
if V_PGI.TauxEuro<=0 then V_PGI.TauxEuro:=6.55957 ;*)
V_PGI.TauxEuro := 1.0 ;
V_PGI.DateDebutEuro:=GetParamSocSecur('SO_DATEDEBUTEURO',iDate1900) ;
{$ENDIF}
// Autres informations
VH^.ExoV8.Code:=Trim(GetParamSocSecur('SO_EXOV8','')) ;
VH^.Cpta[fbGene].Lg:=GetParamSocSecur('SO_LGCPTEGEN',0) ;
Stb:=GetParamSocSecur('SO_BOURREGEN','') ;
if Stb<>'' then VH^.Cpta[fbGene].Cb:=Stb[1] else BEGIN Result:=True ; VH^.Cpta[fbGene].Cb:='0' ; END ;
VH^.Cpta[fbGene].Attente:=GetParamSocSecur('SO_GENATTEND','') ;
VH^.Cpta[fbAux].Lg:=GetParamSocSecur('SO_LGCPTEAUX',0) ;
Stb:=GetParamSocSecur('SO_BOURREAUX','') ;
if Stb<>'' then VH^.Cpta[fbAux].Cb:=Stb[1] else BEGIN Result:=True ; VH^.Cpta[fbAux].Cb:='0' ; END ;
VH^.MontantNegatif:=GetParamSocSecur('SO_MONTANTNEGATIF',False) ;
VH^.DateCloturePer:=GetParamSocSecur('SO_DATECLOTUREPER',iDate1900,True) ;
VH^.DateCloturePro:=GetParamSocSecur('SO_DATECLOTUREPRO',iDate1900,True) ;
VH^.TvaEncSociete:=GetParamSocSecur('SO_TVAENCAISSEMENT','') ;
VH^.EtablisDefaut:=GetParamSocSecur('SO_ETABLISDEFAUT','') ; VH^.EtablisCpta:=GetParamSocSecur('SO_ETABLISCPTA','') ;
VH^.Pays:=GetParamSocSecur('SO_PAYS','') ;
VH^.PaysLocalisation:=QuelPaysLocalisation ; //XVI 24/02/2005
VH^.OuvreBil:=GetParamSocSecur('SO_OUVREBIL','') ;
{Euro}
VH^.TenueEuro:=GetParamSocSecur('SO_TENUEEURO',True) ;
{$IFNDEF EURO}
VH^.CPLienGamme:=GetParamSocSecur('SO_CPLIENGAMME','') ;
{$ENDIF}
//LG 23/09/2002 - suppression de ce test
//if ((VH^.CPLienGamme='S1') or (VH^.CPLienGamme='S3')) then VH^.RegleEquilSais:='TTC' ;
VH^.FBil[1].Deb:=GetParamSocSecur('SO_BILDEB1','') ; VH^.FBil[1].Fin:=GetParamSocSecur('SO_BILFIN1','') ;
VH^.FBil[2].Deb:=GetParamSocSecur('SO_BILDEB2','') ; VH^.FBil[3].Fin:=GetParamSocSecur('SO_BILFIN2','') ;
VH^.FBil[3].Deb:=GetParamSocSecur('SO_BILDEB3','') ; VH^.FBil[3].Fin:=GetParamSocSecur('SO_BILFIN3','') ;
VH^.FBil[4].Deb:=GetParamSocSecur('SO_BILDEB4','') ; VH^.FBil[4].Fin:=GetParamSocSecur('SO_BILFIN4','') ;
VH^.FBil[5].Deb:=GetParamSocSecur('SO_BILDEB5','') ; VH^.FBil[5].Fin:=GetParamSocSecur('SO_BILFIN5','') ;
VH^.FCha[1].Deb:=GetParamSocSecur('SO_CHADEB1','') ; VH^.FCha[1].Fin:=GetParamSocSecur('SO_CHAFIN1','') ;
VH^.FCha[2].Deb:=GetParamSocSecur('SO_CHADEB2','') ; VH^.FCha[2].Fin:=GetParamSocSecur('SO_CHAFIN2','') ;
VH^.FCha[3].Deb:=GetParamSocSecur('SO_CHADEB3','') ; VH^.FCha[3].Fin:=GetParamSocSecur('SO_CHAFIN3','') ;
VH^.FCha[4].Deb:=GetParamSocSecur('SO_CHADEB4','') ; VH^.FCha[4].Fin:=GetParamSocSecur('SO_CHAFIN4','') ;
VH^.FCha[5].Deb:=GetParamSocSecur('SO_CHADEB5','') ; VH^.FCha[5].Fin:=GetParamSocSecur('SO_CHAFIN5','') ;
VH^.FPro[1].Deb:=GetParamSocSecur('SO_PRODEB1','') ; VH^.FPro[1].Fin:=GetParamSocSecur('SO_PROFIN1','') ;
VH^.FPro[2].Deb:=GetParamSocSecur('SO_PRODEB2','') ; VH^.FPro[2].Fin:=GetParamSocSecur('SO_PROFIN2','') ;
VH^.FPro[3].Deb:=GetParamSocSecur('SO_PRODEB3','') ; VH^.FPro[3].Fin:=GetParamSocSecur('SO_PROFIN3','') ;
VH^.FPro[4].Deb:=GetParamSocSecur('SO_PRODEB4','') ; VH^.FPro[4].Fin:=GetParamSocSecur('SO_PROFIN4','') ;
VH^.FPro[5].Deb:=GetParamSocSecur('SO_PRODEB5','') ; VH^.FPro[5].Fin:=GetParamSocSecur('SO_PROFIN5','') ;
VH^.FExt[1].Deb:=GetParamSocSecur('SO_EXTDEB1','') ; VH^.FExt[1].Fin:=GetParamSocSecur('SO_EXTFIN1','') ;
VH^.FExt[2].Deb:=GetParamSocSecur('SO_EXTDEB2','') ; VH^.FExt[2].Fin:=GetParamSocSecur('SO_EXTFIN2','') ;
VH^.VerifCorresp[1,1]:=GetParamSocSecur('SO_CORSGE1',False) ; VH^.VerifCorresp[1,2]:=GetParamSocSecur('SO_CORSGE2',False) ;
VH^.VerifCorresp[2,1]:=GetParamSocSecur('SO_CORSAU1',False) ; VH^.VerifCorresp[2,2]:=GetParamSocSecur('SO_CORSAU2',False) ;
VH^.VerifCorresp[3,1]:=GetParamSocSecur('SO_CORSA11',False) ; VH^.VerifCorresp[3,2]:=GetParamSocSecur('SO_CORSA12',False) ;
VH^.VerifCorresp[4,1]:=GetParamSocSecur('SO_CORSA21',False) ; VH^.VerifCorresp[4,2]:=GetParamSocSecur('SO_CORSA22',False) ;
VH^.VerifCorresp[5,1]:=GetParamSocSecur('SO_CORSA31',False) ; VH^.VerifCorresp[5,2]:=GetParamSocSecur('SO_CORSA32',False) ;
VH^.VerifCorresp[6,1]:=GetParamSocSecur('SO_CORSA41',False) ; VH^.VerifCorresp[6,2]:=GetParamSocSecur('SO_CORSA42',False) ;
VH^.VerifCorresp[7,1]:=GetParamSocSecur('SO_CORSA51',False) ; VH^.VerifCorresp[7,2]:=GetParamSocSecur('SO_CORSA52',False) ;
VH^.VerifCorresp[8,1]:=GetParamSocSecur('SO_CORSBU1',False) ; VH^.VerifCorresp[8,2]:=GetParamSocSecur('SO_CORSBU2',False) ;
VH^.JalCtrlBud:=GetParamSocSecur('SO_JALCTRLBUD','') ;
VH^.DateRevision:=GetParamSocSecur('SO_DATEREVISION',iDate1900) ;
VH^.TauxCoutTiers:=GetParamSocSecur('SO_TAUXCOUTFITIERS',0) ;
VH^.NumPlanRef:=GetParamSocSecur('SO_NUMPLANREF',0) ;
VH^.DefautCli:=GetParamSocSecur('SO_DEFCOLCLI','') ;
VH^.DefautFou:=GetParamSocSecur('SO_DEFCOLFOU','') ; VH^.DefautSal:=GetParamSocSecur('SO_DEFCOLSAL','') ;
VH^.DefautDDivers:=GetParamSocSecur('SO_DEFCOLDDIV',''); VH^.DefautCDivers:=GetParamSocSecur('SO_DEFCOLCDIV','') ;
VH^.DefautDivers:=GetParamSocSecur('SO_DEFCOLDIV','');  ChargeParamTableLibre(GetParamSocSecur('SO_PARAMTABLE',''));
St:=GetParamSocSecur('SO_BOURRELIB','') ; if St<>'' then VH^.BourreLibre:=St[1] ;
VH^.DupSectBud:=GetParamSocSecur('SO_DUPSECTBUD',False) ;
VH^.AttribRIBAuto:=GetParamSocSecur('SO_ATTRIBRIBAUTO',False) ;
VH^.CPControleDoublon:=GetParamSocSecur('SO_CPCONTROLEDOUBLON',False) ;
VH^.CPDateObli:=GetParamSocSecur('SO_CPDATEOBLI',False) ;
VH^.CPCodeAuxiOnly:=GetParamSocSecur('SO_CPCODEAUXIONLY',False) ;
VH^.CPLibreAnaObli:=GetParamSocSecur('SO_CPLIBREANAOBLI',False) ;
VH^.CPChampDoublon:=GetParamSocSecur('SO_CPCHAMPDOUBLON','') ;
VH^.LienPublifi:=GetParamSocSecur('SO_LIENPUBLIFI',False) ; if ctxPCL in V_PGI.PGIContexte then VH^.LienPublifi:=True ;
VH^.CPStatusBarre:=GetParamSocSecur('SO_CPSTATUSBARRE','') ;
// PCL
VH^.CPANAAVANCEE:=GetParamSocSecur('SO_CPANAAVANCEE',False) ;
// Saisie bordereau
VH^.ZSAISIEANAL:=GetParamSocSecur('SO_ZSAISIEANAL',False) ;
VH^.ZSAISIEECHE:=GetParamSocSecur('SO_ZSAISIEECHE',False) ;
VH^.ZGEREANAL:=GetParamSocSecur('SO_ZGEREANAL',False) ;
VH^.ZAUTOSAVE:=GetParamSocSecur('SO_ZAUTOSAVE',False) ;
VH^.ZACTIVEPFU:=GetParamSocSecur('SO_ZACTIVEPFU',False) ;
VH^.ZFOLIOTEMPSREEL:=GetParamSocSecur('SO_ZFOLIOTEMPSREEL',False) ;
VH^.ZSAUVEFOLIOLOCAL:=GetParamSocSecur('SO_ZSAUVEFOLIOLOCAL',False) ;
{#TVAENC}
VH^.OuiTvaEnc:=GetParamSocSecur('SO_OUITVAENC',False) or (VH^.PaysLocalisation=CodeISOES) ; //XMG 31/07/03
if VH^.OuiTvaEnc then
   BEGIN
   VH^.CollCliEnc:=GetParamSocSecur('SO_COLLCLIENC','') ;
   VH^.CollFouEnc:=GetParamSocSecur('SO_COLLFOUENC','') ;
   END Else
   BEGIN
   VH^.CollCliEnc:='' ;
   VH^.CollFouEnc:='' ;
   END ;
  VH^.IsBaseConso:=GetParamSocSecur('SO_CONTACT','')='CONSO' ;
  VH^.PointageJal:=FALSE ;
  VH^.PointageJal:=GetParamSocSecur('SO_POINTAGEJAL',False) ;
  VH^.AnaCroisaxe := GetParamSocSecur('SO_CROISAXE', false);
END ;
{$ENDIF}


{$IFNDEF NOVH}
FUNCTION CHARGESOCIETEHALLEY : Boolean ;
Var Q : TQuery ;
    OkPb : boolean ;
BEGIN
{$IFDEF EURO}
V_PGI.SAV:=FALSE ;
{$ENDIF}
{$IFDEF CHR}
if not VH_GC.ModeDeconnecte then
{$ENDIF CHR}
ChargeSocieteBase ;

{Chargement Table Société ou Paramsoc}
OkPb:=ChargeInfosSociete ;
{$IFDEF EURO}
V_PGI.SAV:=TRUE ;
{$ENDIF}
{Infos devise pivot}
Q:=OpenSQL('Select D_MAXDEBIT, D_MAXCREDIT, D_CPTLETTRDEBIT, D_CPTLETTRCREDIT,D_LIBELLE,D_SYMBOLE from DEVISE Where D_DEVISE="'+V_PGI.DevisePivot+'"',True,-1,'',true);
if Not Q.EOF then
   BEGIN
   VH^.MaxDebitPivot:=Q.Fields[0].AsFloat ; VH^.MaxCreditPivot:=Q.Fields[1].AsFloat ;
   VH^.CptLettrDebit:=Q.Fields[2].AsString ; VH^.CptLettrCredit:=Q.Fields[3].AsString ;
   VH^.LibDevisePivot:=Q.Fields[4].AsString ;
   {$IFNDEF EAGLSERVER}
   V_PGI.SymbolePivot:=Q.Fields[5].AsString ;
   {$ENDIF}
   END ;
Ferme(Q) ;
{Infos devise fongible (pré-initialisé à Devise Pivot)}
(* // SBO 01/07/2007 : FQ20888 --> Il ne faut plus charger le taux de parité des devises fongibles !
if VH^.TenueEuro then
   BEGIN
   Q:=OpenSQL('Select D_DEVISE, D_PARITEEURO,D_LIBELLE, D_SYMBOLE from DEVISE Where D_MONNAIEIN="X" AND D_FONGIBLE="X" AND D_DEVISE<>"'+V_PGI.DevisePivot+'"',True) ;
   if Not Q.EOF then
      BEGIN
      {$IFNDEF EAGLSERVER}
      V_PGI.DeviseFongible:=Q.Fields[0].AsString ;
      V_PGI.TauxEuro:=Q.Fields[1].AsFloat ;
      {$ENDIF}

      VH^.LibDeviseFongible:=Q.Fields[2].AsString ;
      VH^.SymboleFongible:=Q.Fields[3].AsString ;
      END ;
   Ferme(Q) ;
   END ;*)
{Charger les monnaies IN}
VH^.MonnaiesIn:=';' ;
Q:=OpenSQL('Select D_DEVISE from DEVISE Where D_MONNAIEIN="X"',True,-1,'',true);
While Not Q.EOF do
   BEGIN
   VH^.MonnaiesIn:=VH^.MonnaiesIn+Q.Fields[0].AsString+';' ;
   Q.Next ;
   END ;
Ferme(Q) ;
ChargeParamsImmo ;
ChargeParamsCpta ;
{$IFNDEF HVCL}
//  BBI Web services
//{$IFNDEF EAGLSERVER}
ChargeParamsPGI ;
//{$ENDIF}
//  BBI Fin Web services
{$ENDIF}
Result:=Not OkPb ;
END ;
{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/03/2002
Modifié le ... :   /  /
Description .. : LG* suppression de l'utilisation des stringlists
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
Procedure ChargeDescriCPTA ;
Var Q : TQuery ;
BEGIN
   Q:=OpenSQL('SELECT J_JOURNAL, J_NATUREJAL from JOURNAL Where J_NATUREJAL="VTE" or J_NATUREJAL="ACH"',True,-1,'',true);
   if Not Q.EOF then
      BEGIN
      VH^.ListeNatJal:=';' ;
      While Not Q.EOF do
         BEGIN
         VH^.ListeNatJal:=VH^.ListeNatJal+Q.Fields[0].AsString+';' ;
         Q.Next ;
         END ;
      END ;
   Ferme(Q) ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
PROCEDURE CHARGESOUSPLANAXE ;
Var fb : TFichierBase ;
    Q,Q1 : TQuery ;
    QAxe,QCode,QLib : String ;
    QDeb,QLong : Integer ;
    TabI : Array[fbAxe1..fbAxe5] Of Integer ;
    i : Integer ;
BEGIN
For fb:=fbAxe1 to fbAxe5 Do For i:=1 To MaxSousPlan Do If VH^.SousPlanAxe[fb,i].ListeSP<>NIL Then
  BEGIN
  VH^.SousPlanAxe[fb,i].ListeSP.Clear ; VH^.SousPlanAxe[fb,i].ListeSP.Free ;
  END ;
Fillchar(VH^.SousPlanAxe,SizeOf(VH^.SousPlanAxe),#0) ;
Fillchar(TabI,SizeOf(TabI),#0) ;
//GP 20/04/98 Q:=OpenSQL('Select * From STRUCRSE Order By SS_AXE, SS_SOUSSECTION',True) ;
Q:=OpenSQL('Select * From STRUCRSE Order By SS_AXE, SS_DEBUT',True, -1, 'STRUCRSE',true) ;
While Not Q.Eof Do
  BEGIN
  QAxe:=Q.FindField('SS_AXE').AsString ;
  QCode:=Q.FindField('SS_SOUSSECTION').AsString ;
  QLib:=Q.FindField('SS_LIBELLE').AsString ;
  QDeb:=Q.FindField('SS_DEBUT').AsInteger ;
  QLong:=Q.FindField('SS_LONGUEUR').AsInteger ;
  fb:=AxeTofb(QAxe) ;
  Inc(TabI[fb]) ;
  If TabI[fb]<=MaxSousPlan Then
    BEGIN
    VH^.SousPlanAxe[fb,TabI[fb]].Code:=QCode ; VH^.SousPlanAxe[fb,TabI[fb]].Lib:=QLib ;
    VH^.SousPlanAxe[fb,TabI[fb]].Debut:=QDeb ; VH^.SousPlanAxe[fb,TabI[fb]].Longueur:=QLong ;
    VH^.SousPlanAxe[fb,TabI[fb]].ListeSP:= HTStringList.Create ;
    Q1:=OpenSQL('Select * from SSSTRUCR WHERE PS_AXE="'+QAxe+'" AND PS_SOUSSECTION="'+QCode+'" ORDER BY PS_CODE',TRUE,-1,'SSSTRUCR',true) ;
    While Not Q1.Eof Do
      BEGIN
      VH^.SousPlanAxe[fb,TabI[fb]].ListeSP.Add(Q1.FindField('PS_CODE').ASString+';'+Q1.FindField('PS_LIBELLE').ASString+';') ;
      Q1.Next ;
      END ;
    Ferme(Q1) ;
    END ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
PROCEDURE CHARGECATBUD ;
Var fb : TFichierBase ;
    Q : TQuery ;
    StCode,StLib,StLibre,StS,StJ,St : String ;
    i,k,l : Integer ;
BEGIN
Fillchar(VH^.CatBud,SizeOf(VH^.CatBud),#0) ;
Q:=OpenSQL('Select * From CHOIXCOD WHERE CC_TYPE="CJB"',TRUE,-1,'CHOIXCOD',true) ; k:=1 ;
While (Not Q.Eof) And (k<=MaxCatBud) Do
  BEGIN
  StCode:=Q.FindField('CC_CODE').AsString ;
  StLib:=Q.FindField('CC_LIBELLE').AsString ;
  StLibre:=Q.FindField('CC_LIBRE').AsString ;
  fb:=AxeTofb(Q.FindField('CC_ABREGE').AsString) ;
  i:=Pos('/',StLibre) ;
  If i>0 Then
     BEGIN
     StS:=Copy(StLibre,1,i-1) ; StJ:=Copy(StLibre,i+1,Length(StLibre)-i) ;
     VH^.CatBud[k].fb:=fb ; VH^.CatBud[k].Code:=StCode ; VH^.CatBud[k].Lib:=StLib ;
     l:=1 ;
     While StS<>'' Do
       BEGIN
       St:=ReadTokenSt(StS) ;
       If St<>'' Then BEGIN VH^.CatBud[k].SurSect[l]:=St ; Inc(l) ; END ;
       END ;
     l:=1 ;
     While StJ<>'' Do
       BEGIN
       St:=ReadTokenSt(StJ) ;
       If St<>'' Then BEGIN VH^.CatBud[k].SurJal[l]:=St ; Inc(l) ; END ;
       END ;
     Inc(k) ;
     END ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;
{$ENDIF}

Function QuelleCatBud(CodeCat : String) : TUneCatBud ;
Var i,j : Integer ;
    Cat : TUneCatBud ;
BEGIN
j:=0 ; Fillchar(Cat,SizeOf(Cat),#0) ;
For i:=1 To MaxCatBud Do If GetCatBud(i).Code=CodeCat Then j:=i ; { Retrouve la catégorie }
If j<>0 Then Cat:=GetCatBud(j);
Result:=Cat ;
END ;

Function SousPlanCat(CodeCat : String ; SurJal : Boolean) : TSousPlanCat ;
Var i,j,l,Ind : Integer ;
    Cat : TUneCatBud ;
    SPC : TSousPlanCat ;
    St : String ;
BEGIN
Fillchar(SPC,SizeOf(SPC),#0) ; Ind:=0 ;
Cat:=QuelleCatBud(CodeCat) ;
If Cat.Code<>'' Then
   BEGIN
   For l:=1 To MaxSousPlan Do
     BEGIN
     If SurJal Then St:=Cat.SurJal[l] Else St:=Cat.SurSect[l] ;
     { Prend le l° code sous plan de la catégorie trouvée }
     If St<>'' Then
        BEGIN
        j:=0 ;
        //For i:=1 To MaxSousPlan Do If VH^.SousPlanAxe[Cat.fb,i].Code=St Then j:=i ;
        For i:=1 To MaxSousPlan Do If GetSousPlanAxe(Cat.fb,i).Code=St Then j:=i ;
        { retrouve les caractéristiques du sous plan concerné }
        If j<>0 Then
           BEGIN
           Inc(Ind) ;
           SPC[Ind] := GetSousPlanAxe(Cat.fb,j);
           END ;
        END ;
     END ;
   END ;
Result:=SPC ;
END ;

{$IFNDEF NOVH}
Procedure ChargeHalleyUser ;
Var QG : TQuery ;
BEGIN
QG:=OpenSQL('SELECT UG_NIVEAUACCES, UG_NUMERO, UG_LIBELLE, UG_CONFIDENTIEL, UG_MONTANTMIN, UG_MONTANTMAX, UG_JALAUTORISES, UG_LANGUE, UG_PERSO FROM USERGRP WHERE UG_GROUPE="'+V_PGI.Groupe+'"',TRUE, -1, 'USERGRP',true) ;
if Not QG.EOF then
   BEGIN
   VH^.GrpMontantMin:=QG.FindField('UG_MONTANTMIN').AsFloat ;
   VH^.GrpMontantMax:=QG.FindField('UG_MONTANTMAX').AsFloat ;
   VH^.JalAutorises:=QG.FindField('UG_JALAUTORISES').AsString ;
   // CA - 18/07/2007 - FQ 20027 - Cas du << Tous >> enregistré dans la base
   if (Pos('<<',VH^.JalAutorises)>0) then VH^.JalAutorises:= '';
   if VH^.JalAutorises='-' then VH^.JalAutorises := ''; // 13569
   if VH^.JalAutorises<>'' then
      BEGIN
      if VH^.JalAutorises[1]<>';' then VH^.JalAutorises:=';'+VH^.JalAutorises ;
      if VH^.JalAutorises[Length(VH^.JalAutorises)]<>';' then VH^.JalAutorises:=VH^.JalAutorises+';' ;
      END ;
   if (VH^.GrpMontantMax<=VH^.GrpMontantMin) then
     begin VH^.GrpMontantMin:=0 ; VH^.GrpMontantMax:=999999999999.0 ; end ;
   END else
   BEGIN
   VH^.GrpMontantMin:=0 ; VH^.GrpMontantMax:=999999999999.0 ; VH^.JalAutorises:='' ;
   END ;
Ferme(QG) ;
END;
{$ENDIF}

{$IFNDEF NOVH}
Procedure ChargeProfilUser ;
Var Q : TQuery ;
    TT : String ;
    ind : tProfilTraitement ;
    OkOk,BPRO : Boolean ;
BEGIN
Fillchar(VH^.ProfilUserC,SizeOf(VH^.ProfilUserC),#0) ;
{$IFDEF ALLVERSION}
If VH^.VNUMSOC<500 Then Exit ;
BPRO := false;
{$ENDIF}
VH^.ProfilCCMPExiste[prClient]:=FALSE ;
VH^.ProfilCCMPExiste[prFournisseur]:=FALSE ;
VH^.ProfilUserC[prEtablissement].ForceDepot := False; // JTR - Force le dépot
{$IFNDEF EAGL}
if TableExiste('BPROFILUSERC') then
begin
  Q:=OpenSQL('SELECT * FROM CPPROFILUSERC LEFT JOIN BPROFILUSERC ON BPU_USER=CPU_USER WHERE CPU_USERGRP="'+V_PGI.Groupe+'" AND CPU_USER="'+V_PGI.User+'" ',TRUE, -1 ,'CPPROFILUSERC',true) ;
  BPRO := true;
end else
begin
  Q:=OpenSQL('SELECT * FROM CPPROFILUSERC WHERE CPU_USERGRP="'+V_PGI.Groupe+'" AND CPU_USER="'+V_PGI.User+'" ',TRUE, -1 ,'CPPROFILUSERC',true) ;
end;
{$ELSE}
Q:=OpenSQL('SELECT * FROM CPPROFILUSERC WHERE CPU_USERGRP="'+V_PGI.Groupe+'" AND CPU_USER="'+V_PGI.User+'" ',TRUE,-1,'',true);
{$ENDIF}
While Not Q.Eof Do
  BEGIN
  TT:=Q.FindField('CPU_TYPE').AsString ; Okok:=TRUE ; Ind:=prAucun ;
  If TT='CLI' Then Ind:=prClient Else If TT='FOU' Then Ind:=prFournisseur Else
     if TT='ETA' then Ind:=prEtablissement else OkOk:=FALSE ;
  If OkOk Then
    BEGIN
    VH^.ProfilUserC[Ind].Compte1:=Q.FindField('CPU_COMPTE1').AsString ;
    VH^.ProfilUserC[Ind].Compte2:=Q.FindField('CPU_COMPTE2').AsString ;
    VH^.ProfilUserC[Ind].Exclusion1:=Q.FindField('CPU_EXCLUSION1').AsString ;
    VH^.ProfilUserC[Ind].Exclusion2:=Q.FindField('CPU_EXCLUSION2').AsString ;
    VH^.ProfilUserC[Ind].TableLibre1:=Q.FindField('CPU_TABLELIBRE1').AsString ;
    VH^.ProfilUserC[Ind].TableLibre2:=Q.FindField('CPU_TABLELIBRE2').AsString ;
    VH^.ProfilUserC[Ind].ForceEtab:=(Q.FindField('CPU_FORCEETAB').AsString='X') ;
    VH^.ProfilUserC[Ind].Domaine:=Q.FindField('CPU_DOMAINE').AsString ;
    VH^.ProfilUserC[Ind].ForceDomaine:=(Q.FindField('CPU_FORCEDOMAINE').AsString='X') ;
    VH^.ProfilUserC[Ind].Etablissement:=Q.FindField('CPU_ETABLISSEMENT').AsString ;
    if BPRO then
    begin
      VH^.ProfilUserC[Ind].responsable:=Q.FindField('BPU_RESPONSABLE').AsString ;
      VH^.ProfilUserC[Ind].ForceResponsable:=(Q.FindField('BPU_FORCERESP').AsString='X') ;
    end else
    begin
      VH^.ProfilUserC[Ind].responsable := '';
      VH^.ProfilUserC[Ind].ForceResponsable := false;
    end;
    // JTR Dépôt par défaut et PCP
    if V_PGI.NumVersionBase >= 631 then
    begin
      if Ind = prEtablissement then
      begin
        VH^.ProfilUserC[Ind].Depot:=Q.FindField('CPU_DEPOT').AsString ;
        VH^.ProfilUserC[Ind].ForceDepot:=(Q.FindField('CPU_FORCEDEPOT').AsString='X') ; // JTR - Force le dépot
        VH^.ProfilUserC[Ind].PCPVente:=Q.FindField('CPU_PCPVENTE').AsString ;
        VH^.ProfilUserC[Ind].PCPAchat:=Q.FindField('CPU_PCPACHAT').AsString ;
      end else
      begin
        VH^.ProfilUserC[Ind].Depot:='';
        VH^.ProfilUserC[Ind].PCPVente:='';
        VH^.ProfilUserC[Ind].PCPAchat:='';
      end;
    end;
    // Fin JTR
    VH^.ProfilUserC[Ind].ProfilOk:=TRUE ;
    VH^.ProfilCCMPExiste[Ind]:=TRUE ;
    END ;
  Q.Next ;
  END ;
Ferme(Q) ;

// Restriction par groupe non utilisable en compta : SBO 05/06/2005
if (not (ctxCompta in V_PGI.PGIContexte)) and (not (ctxTreso in V_PGI.PGIContexte)) then
  // JTR Dépôt par défaut
  if V_PGI.NumVersionBase >= 631 then
  begin
    if (VH^.ProfilUserC[prEtablissement].Depot = '') or (VH^.ProfilUserC[prEtablissement].Etablissement = '') then
    begin
        Q := OpenSQL('SELECT * FROM CPPROFILUSERC WHERE CPU_USERGRP="'+V_PGI.Groupe+'" AND CPU_USER="..." ',TRUE, -1 ,'CPPROFILUSERC',true) ;
        if not Q.Eof then
        begin
          if VH^.ProfilUserC[prEtablissement].Depot = '' then
          begin
            VH^.ProfilUserC[prEtablissement].Depot := Q.FindField('CPU_DEPOT').AsString;
            VH^.ProfilUserC[prEtablissement].ForceDepot := (Q.FindField('CPU_FORCEDEPOT').AsString='X') ; // JTR - Force le dépot
          end;
          if VH^.ProfilUserC[prEtablissement].Etablissement = '' then
          begin
            VH^.ProfilUserC[prEtablissement].Etablissement := Q.FindField('CPU_ETABLISSEMENT').AsString;
            // VH^.ProfilUserC[Ind].ForceEtab:=(Q.FindField('CPU_FORCEETAB').AsString='X') ; // JTR - 30/06/2004
            VH^.ProfilUserC[prEtablissement].ForceEtab:=(Q.FindField('CPU_FORCEETAB').AsString='X') ;
          end;
          if VH^.ProfilUserC[prEtablissement].Domaine = '' then
          begin
            VH^.ProfilUserC[prEtablissement].Domaine := Q.FindField('CPU_DOMAINE').AsString;
            VH^.ProfilUserC[prEtablissement].ForceDomaine := (Q.FindField('CPU_FORCEDOMAINE').AsString='X') ; // JTR - Force le dépot
          end;
        end;
      Ferme(Q);
    end;
  end;

// Fin JTR
If Not VH^.ProfilCCMPExiste[prClient] Then
  BEGIN
  Q:=OpenSQL('SELECT * FROM CPPROFILUSERC WHERE CPU_TYPE="CLI" ',TRUE,-1, 'CPPROFILUSERC',true) ;
  If Not Q.Eof Then VH^.ProfilCCMPExiste[prClient]:=TRUE ;
  Ferme(Q) ;
  END ;
If Not VH^.ProfilCCMPExiste[prFournisseur] Then
  BEGIN
  Q:=OpenSQL('SELECT * FROM CPPROFILUSERC WHERE CPU_TYPE="FOU" ',TRUE,-1,'CPPROFILUSERC',true) ;
  If Not Q.Eof Then VH^.ProfilCCMPExiste[prFournisseur]:=TRUE ;
  Ferme(Q) ;
  END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Function WhereProfilCpt(Cpt,Exclu : String ; fb : tFichierBase ; OkTableLibre : Boolean) : String ;
Var Where , Sauf: String ;
BEGIN
Result:='' ;
{$IFDEF ALLVERSION}
If VH^.VNUMSOC<500 Then Exit ;
{$ENDIF}
If Cpt<>'' Then
  BEGIN
  If Trim(Cpt)='TOUS' Then
    BEGIN
    Case fb Of
      fbGene : Where:=' AND (G_GENERAL<>"'+W_W+'") ' ;
      fbAux : Where:='AND (T_AUXILIAIRE<>"'+W_W+'") ' ;
      END ;
    END Else
    BEGIN
    Where:=AnalyseCompte(Cpt,fb,False,OkTableLibre,TRUE) ;
    END ;
  Sauf:=AnalyseCompte(Exclu,fb,TRUE,FALSE,TRUE) ;
  If Where<>'' Then Result:=Result+' AND '+Where ;
  If Not OkTableLibre Then If Sauf<>'' Then Result:=Result+' AND '+Sauf ;
  END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Function WhereProfilUser(Q : TQuery ; pf : tProfilTraitement) : String ;
//Var Where,Sauf,St : String ;
BEGIN
Result:='' ; If Q=NIL Then Exit ;
{$IFDEF ALLVERSION}
If VH^.VNUMSOC<500 Then Exit ;
{$ENDIF}
Result:=Result+WhereProfilCpt(VH^.ProfilUserC[pf].Compte1,VH^.ProfilUserC[pf].Exclusion1,fbGene,FALSE) ;
Result:=Result+WhereProfilCpt(VH^.ProfilUserC[pf].Compte2,VH^.ProfilUserC[pf].Exclusion2,fbAux,FALSE) ;
Result:=Result+WhereProfilCpt(VH^.ProfilUserC[pf].TableLibre1,'',fbGene,TRUE) ;
Result:=Result+WhereProfilCpt(VH^.ProfilUserC[pf].TableLibre2,'',fbAux,TRUE) ;
If Result<>'' Then Delete(Result,1,4) ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Function EtabForce : String ;
Var Etab : String ;
    Forcer : boolean ;
BEGIN
Result:='' ;
{$IFDEF ALLVERSION}
If VH^.VNUMSOC<500 Then Exit ;
{$ENDIF}
Etab:=VH^.ProfilUserC[prEtablissement].Etablissement ; if Etab='' then Exit ;
Forcer:=VH^.ProfilUserC[prEtablissement].ForceEtab ; if Not Forcer then Exit ;
Result:=Etab ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure PositionneResponsableUser (LeResp : TControl; IgnoreForce : Boolean = False; IgnoreVisible : Boolean = False; IgnoreDisabled : boolean=false) ;
Var Resp : String ;
    Forcer : Boolean ;
begin
  if LeResp=Nil then Exit ;
//  if Not VH^.EtablisCpta then Exit ;
  if (not IgnoreVisible) and (Not LeResp.Visible) then Exit ;
  Resp:=VH^.ProfilUserC[prEtablissement].responsable ;
  if Resp='' then Exit ;
  Forcer:=VH^.ProfilUserC[prEtablissement].ForceResponsable ;
  If LeResp is ThValComboBox then
   begin
   if ThValCOmboBox(LeResp).Values.IndexOf(Resp)<0 then Exit ;
   if ((not IgnoreDisabled) and (Not LeResp.Enabled) and (ThValCOmboBox(LeResp).Value<>Resp)) then Exit ;
   ThValCOmboBox(LeResp).Value:=Resp ;
   end
  else  If LeResp is ThMultiValComboBox then
   begin
   if ThMultiValCOmboBox(LeResp).Values.IndexOf(Resp)<0 then Exit ;
   if ((not IgnoreDisabled) and (Not LeResp.Enabled) and (ThMultiValCOmboBox(LeResp).Text<>Resp)) then Exit ;
   ThMultiValCOmboBox(Resp).Text:=Resp ;
   end
  else  If LeResp is ThEdit then
   begin
   if ((not IgnoreDisabled) and (Not LeResp.Enabled) and (THEdit(LeResp).Text<>Resp)) then Exit ;
   ThEdit(LeResp).Text:=Resp ;
   end;
  if ((Forcer) and (Not IgnoreForce)) then LeResp.Enabled:=False ;
end;

Procedure PositionneEtabUser ( LeEtab : TControl ; IgnoreForce : Boolean = False; IgnoreVisible : Boolean = False; IgnoreDisabled : boolean=false) ;
Var Etab : String ;
    Forcer : Boolean ;
BEGIN
  if LeEtab=Nil then Exit ;
{$IFDEF ALLVERSION}
  If VH^.VNUMSOC<500 Then Exit ;
{$ENDIF}
  if Not VH^.EtablisCpta then Exit ;
  if (not IgnoreVisible) and (Not LeEtab.Visible) then Exit ;
  Etab:=VH^.ProfilUserC[prEtablissement].Etablissement ; if Etab='' then Exit ;
  Forcer:=VH^.ProfilUserC[prEtablissement].ForceEtab ;
  If LeEtab is ThValComboBox then
   begin
   if ThValCOmboBox(LeEtab).Values.IndexOf(Etab)<0 then Exit ;
   if ((not IgnoreDisabled) and (Not LeEtab.Enabled) and (ThValCOmboBox(LeEtab).Value<>Etab)) then Exit ;
   ThValCOmboBox(LeEtab).Value:=Etab ;
   end
  else  If LeEtab is ThMultiValComboBox then
   begin
   if ThMultiValCOmboBox(LeEtab).Values.IndexOf(Etab)<0 then Exit ;
   if ((not IgnoreDisabled) and (Not LeEtab.Enabled) and (ThMultiValCOmboBox(LeEtab).Text<>Etab)) then Exit ;
   ThMultiValCOmboBox(LeEtab).Text:=Etab ;
   end;
  if ((Forcer) and (Not IgnoreForce)) then LeEtab.Enabled:=False ;
END ;
{$ENDIF}

// JTR - Idem que pour établissement
{$IFNDEF NOVH}
Procedure PositionneDepotUser (LeDepot : TControl; IgnoreForce : Boolean = False; IgnoreVisible : boolean=false; IgnoreDisabled : boolean=false);
Var Depot : String ;
    Forcer : Boolean ;
begin
  if LeDepot = Nil then Exit ;
  {$IFDEF ALLVERSION}
  if VH^.VNUMSOC < 500 Then Exit ;
  {$ENDIF}
  if (not IgnoreVisible) and (not LeDepot.Visible) then Exit ;
  Depot := VH^.ProfilUserC[prEtablissement].Depot;
  if Depot = '' then Exit;
  Forcer := VH^.ProfilUserC[prEtablissement].ForceDepot;
  if LeDepot is THValComboBox then
  begin
    if ThValCOmboBox(LeDepot).Values.IndexOf(Depot) < 0 then Exit ;
    if ((not IgnoreDisabled) and (Not LeDepot.Enabled) and (ThValCOmboBox(LeDepot).Value <> Depot)) then Exit ;
    ThValCOmboBox(LeDepot).Value := Depot ;
  end else
  if LeDepot is THMultiValComboBox then
  begin
    if THMultiValComboBox(LeDepot).Values.IndexOf(Depot) < 0 then Exit ;
    if ((not IgnoreDisabled) and (Not LeDepot.Enabled) and (THMultiValComboBox(LeDepot).Value <> Depot)) then Exit ;
    THMultiValComboBox(LeDepot).Value := Depot;
  end;
  if ((Forcer) and (Not IgnoreForce)) then LeDepot.Enabled:=False;
end;
{$ENDIF}

{$IFNDEF NOVH}
Procedure PositionneDomaineUser ( LeDom : THValComboBox ; IgnoreForce : Boolean = False; IgnoreVisible : Boolean = False; IgnoreDisabled : boolean=false) ;
Var sDom : String ;
BEGIN
  if LeDom=Nil then Exit ;
  if (Not LeDom.Visible) and (not IgnoreVisible) then Exit ;
  {$IFDEF ALLVERSION}
  If VH^.VNUMSOC<500 Then Exit ;
  {$ENDIF}
  sDom:=VH^.ProfilUserC[prEtablissement].Domaine ;
  if sDom='' then Exit ;
  if THValCOmboBox(LeDom).Values.IndexOf(sDom) < 0 then Exit ;
  if ((not IgnoreDisabled) and (Not LeDom.Enabled) and (ThValCOmboBox(LeDom).Value<>sDom) and (LeDom.Value<>'')) then Exit ;
  THValCOmboBox(LeDom).Value := sDom ;
  if ((VH^.ProfilUserC[prEtablissement].ForceDomaine) and (Not IgnoreForce)) then
    THValCOmboBox(LeDom).Enabled := False ;
END ;
{$ENDIF}

{$IFNDEF EAGLSERVER}
{$IFNDEF NOVH}
Procedure AGLPositionneDomaineUser ( Parms : array of variant ; nb : integer ) ;
Var FF : TForm ;
    NomDom : String ;
    CC     : THValComboBox ;
BEGIN
{$IFDEF ALLVERSION}
If VH^.VNUMSOC<500 Then Exit ;
{$ENDIF}
FF:=TForm(Longint(Parms[0])) ; NomDom:=String(Parms[1]) ;
CC:=THValComboBox(FF.FindComponent(NomDom)) ; if CC=Nil then Exit ;
PositionneDomaineUser(CC) ;
END ;
{$ENDIF}
{$ENDIF}

Function ParamEuroOk : Boolean ;
BEGIN
(*Result:=(V_PGI.DateDebutEuro>=encodeDate(1999,01,01)) And (Trim(VH^.EccEuroDebit)<>'') And
        (Trim(VH^.EccEuroCredit)<>'') And (Trim(VH^.RegleEquilSais)<>'') ;*)
  Result := true;
END ;

{$IFNDEF NOVH}
Procedure ChargeStructureUnique ;
Var QLoc,QLoc1 : TQuery ;
    fb         : TFichierBase ;
BEGIN
QLoc:=OpenSql('Select * From STRUCRSE Where SS_CONTROLE="X" ',True, -1, 'STRUCRSE',true) ;
While Not QLoc.Eof do
     BEGIN
     fb:=AxeToFb(QLoc.FindField('SS_AXE').AsString) ;
     VH^.Cpta[fb].UneStruc.CodeSp:=QLoc.FindField('SS_SOUSSECTION').AsString ;
     VH^.Cpta[fb].UneStruc.Debu:=QLoc.FindField('SS_DEBUT').AsInteger ;
     VH^.Cpta[fb].UneStruc.Long:=QLoc.FindField('SS_LONGUEUR').AsInteger ;
     QLoc1:=OpenSql('Select PS_CODE From SSSTRUCR Where PS_AXE="'+QLoc.FindField('SS_AXE').AsString+'" And '+
                    'PS_SOUSSECTION="'+QLoc.FindField('SS_SOUSSECTION').AsString+'" ',True,-1,'',true) ;
     VH^.Cpta[fb].UneStruc.QuelCarac:=QLoc1.Fields[0].AsString ; Ferme(QLoc1) ;
     QLoc.Next ;
     END ;
Ferme(QLoc) ;
END ;
{$ENDIF}

procedure CheckLiaisonComptableY2;
var QQ: TQuery;
begin
  // si pas de partage de la table CSOUSSECTION(Y2) alors pas de liaison avec Y2 via deshare
  VH^.LiaisonY2ViaShare := false;
  if not ExisteSQL('SELECT 1 FROM DESHARE WHERE DS_NOMTABLE="CSECTION"') then Exit;
  QQ := OpenSQL('SELECT COUNT(*) AS NBR FROM CSECTION',True,1,'',true);
  if not QQ.Eof then
  begin
    VH^.LiaisonY2ViaShare := (QQ.Fields[0].AsInteger >0);
  end;
  ferme (QQ);
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 28/10/2002
Modifié le ... : 05/04/2005
Description .. : - 28/10/2002 - suppression du forcage des param. de
Suite ........ : lettrage ( vu avec O Guillaud )
Suite ........ : - 05/04/2005 - LG  - Ajout du ctsExercice
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
FUNCTION CHARGEMAGHALLEY : Boolean ;
Var fb     : TFichierBase ;
    Q      : TQuery ;
    PbExo  : boolean ;
    NumAxe : integer ;
    StB,StY: String ;
    stMessage : string ;
    {$IFNDEF EAGLSERVER}
      OkPb,PbBoot: boolean;
      {$IFDEF COMPTA}
       ExoEnCours: TExoDate;
      {$ENDIF COMPTA}
    {$ENDIF NOVH}
    {$IFNDEF EAGLSERVER}
    StrDossier : String ;
    StrParam : String ;
    {$ENDIF EAGLSERVER}
    oT : TOB;
		VersionRef : string;
BEGIN

//  TSendEntryY2.RecupParamCptaFromWS;

{$IFNDEF EAGLSERVER}
// MB : Nécessite AGL 20.21 avec unité FiltresDonnees
// Chargement de la surcharge des groupes par utilisateur et données (NoDossier).
if ctxPCL in V_PGI.PGIContexte then
begin
  // MB Le 29/11/2007 Gestion des programmes lancés en DBCOM=TRUE.
  StrDossier := V_PGI.Nodossier ;
  if ( StrDossier = '000000' ) then
  begin
     StrParam := '' ;
     If PGIApp <> Nil Then
        StrParam := PGIApp.GetParam('General', 'CURNODOSSIER','') ;
     if StrParam <> '' then
        StrDossier := StrParam ;
     // PgiBox('Dossier pour surcharge = '+ StrDossier) ;
  end;
  if V_PGI.GrpsDelegues <> '' then
     V_PGI.GrpsDelegues := V_PGI.GrpsDelegues + ';' ;
  V_PGI.GrpsDelegues := V_PGI.GrpsDelegues + GetProfilsFromDonnees('GROUPECONF',V_PGI.User, V_PGI.Nodossier) ;
  If ( StrDossier <> V_PGI.Nodossier ) then
  begin
    if V_PGI.GrpsDelegues <> '' then
       V_PGI.GrpsDelegues := V_PGI.GrpsDelegues + ';' ;
    V_PGI.GrpsDelegues := V_PGI.GrpsDelegues + GetProfilsFromDonnees('GROUPECONF',V_PGI.User, StrDossier) ;
  end;
  AglLoadGrpsDelegues ;
  // Aussi sur le Nodossier = 000000
end;
{$ENDIF EAGLSERVER}

VH^.ExoV8.Code:='' ; VH^.ExoV8.Deb:=0 ; VH^.ExoV8.Fin:=0 ;
{$IFNDEF EAGLSERVER}
OkPb:=False ;
{$ENDIF NOVH}
PbExo:=False ;
VH^.VNUMSOC:=500 ;
{$IFDEF ALLVERSION}
Q:=OPENSQL('SELECT SO_VERSIONBASE FROM SOCIETE ',TRUE,-1,'',true) ;
If Not Q.Eof Then VH^.VNUMSOC:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
{$ENDIF}
V_PGI.CergEF:=FALSE ;
{$IFNDEF EAGLSERVER}
PbBoot:=False ;
//
CheckLiaisonComptableY2; // LSLSLSLS
//
{$ENDIF EAGLSERVER}
if ((V_PGI.User='000') and (V_PGI.UserName='00000000') and (V_PGI.PassWord='000')) then
   BEGIN
   Result:=TRUE ; if V_PGI.Halley then Exit ;
   END else Result:=TRUE ;

   Q:=OPENSQL('SELECT SO_VERSIONBASE FROM SOCIETE ',TRUE,-1,'',true) ;
   If Not Q.Eof Then VH^.VNUMSOC:=Q.Fields[0].AsInteger ;
   Ferme(Q) ;
   if v_pgi.NumversionBase>VH^.VNUMSOC then
   begin
      stMessage := 'Cette société (' + inttostr(VH^.VNUMSOC) + ') nécessite une mise à jour de structure.';
      stMessage := stMessage +#13 +'Vous devez lancer l''utilitaire Administration de base PGI.';
      stMessage := stMessage +#13 ;
      stMessage := stMessage +#13 +'Contacter votre administrateur.';
      PGIInfo(stMessage,'Erreur de version de base');
      result := false ;
   end ;

if not result then exit;
  if (not ISUtilisable) then
  begin
     stMessage := 'Base de données momentanément indisponible (mise à niveau en cours) ';
     stMessage := stMessage +#13#10 +'Veuillez réessayer ultérieurement';
     PGIInfo(stMessage,'Information de connection');
     result := false;
  end;
  if not result then exit;
{$IFNDEF MAJCEGID}
  versionRef := '';
{$IFDEF EAGLCLIENT}
  // recup version base ref
  oT := AppServer.Request('PluginBTPS.GetVersionRef', 'INFO', nil, '', '');
  if assigned(Ot) and (Ot.FieldExists('VERSIONREF') ) then
  begin
    VersionRef := Ot.GetString ('VERSIONREF');
    OT.free;
  end;
{$ELSE}
	versionRef := BTPGetVersionRef;
{$ENDIF}
  if versionRef = '' then
  begin
     stMessage := 'Merci de contacter votre administrateur';
     stMessage := stMessage +#13#10 +'La version de la base de réference n''est pas disponible.';
     PGIInfo(stMessage,'INFORMATION ADMINISTRATION');
     result := false;
  end;
  if not result then exit;
{$IFNDEF PGIMAJVER}
  if not Okversion (VersionRef) then
  begin
     stMessage := 'Merci de contacter votre administrateur';
     stMessage := stMessage +#13#10 +'afin de mettre à niveau votre base de données';
     PGIInfo(stMessage,'INFORMATION VERSION');
     result := false;
  end;
{$ENDIF}
{$IFDEF PGIMAJVER}
  if Result then
  begin
    if not DelphiRunning then
    begin
      // Controle que le programme est en bonne version pour se connecter a la base de donnée
      if BTPGetVersionRef <> NumRefObligatoire then
      begin
        stMessage := 'Merci de contacter votre administrateur';
        stMessage := stMessage +#13#10 +'afin de mettre la version de produit correcte sur votre poste';
        PGIInfo(stMessage,'INFORMATION VERSION DE PROGRAMME');
        Result := false;
      end;
    end;
  end;
{$ENDIF}
{$ENDIF} // CEGID

if Not Result then exit ;
Try
 BeginTrans ;
 InitMove(23,'Initialisation ...') ;
 MoveCur(FALSE) ; ChargeTVATPF ;
 MoveCur(FALSE) ;
// JTR Dépôt par défaut - Déplacement de ChargeProfilUser
{$IFNDEF SPEC302}
 {$IFNDEF SPEC350}
 ChargeProfilUser ;
 {$ENDIF}
{$ENDIF}
// Fin JTR
{$IFNDEF NOVH}
 MoveCur(FALSE) ; if not ChargeSocieteHalley then
 {$IFNDEF EAGLSERVER}
 OkPb:=True
 {$ENDIF}
 ;
{$ENDIF NOVH}
 MoveCur(FALSE) ;
 {$IFDEF SPEC302}
 ExecuteSQL('UPDATE SOCIETE SET SO_DATEDERNENTREE="'+USDATETIME(Date)+'"') ;
 MAJMENU320 ;
 {$ELSE}
 {$IFNDEF EAGLCLIENT}
 SetParamsoc('SO_DATEDERNENTREE',Date) ;
 {$ENDIF}
 {$ENDIF}
 MoveCur(FALSE) ; NumAxe:=0 ;
 Q:=OpenSQL('SELECT * FROM AXE ORDER BY X_AXE',TRUE, -1, 'AXE',true) ;
 While Not Q.EOF do
    BEGIN
    Inc(NumAxe) ; fb:=AxeToFb(Q.FindField('X_AXE').AsString) ;
    VH^.Cpta[fb].Lg:=Q.FindField('X_LONGSECTION').AsInteger ;
    Stb:=Q.FindField('X_BOURREANA').AsString ;
    if Stb<>'' then
       BEGIN
       VH^.Cpta[fb].Cb:=Stb[1] ;
       END else
       BEGIN
{$IFNDEF EAGLSERVER}
        OkPb:=True ;
{$ENDIF EAGLSERVER}
        VH^.Cpta[fb].Cb:='0' ;
       END ;
 {$IFNDEF SPEC302}
    VH^.SaisieTranche[NumAxe]:=(Q.FindField('X_SAISIETRANCHE').AsString='X') ;
 {$ENDIF}
    VH^.Cpta[fb].Chantier:=Q.FindField('X_CHANTIER').AsString='X' ;
    VH^.Cpta[fb].Structure:=Q.FindField('X_STRUCTURE').AsString='X' ;
    VH^.Cpta[fb].Attente:=Q.FindField('X_SECTIONATTENTE').AsString ;
    VH^.Cpta[fb].AxGenAttente:=Q.FindField('X_GENEATTENTE').AsString ;
    Q.Next ;
    END ;
 Ferme(Q) ;
 MoveCur(FALSE) ;
 MoveCur(False) ;
 {$IFNDEF PGIIMMO}
 ChargeStructureUnique ;
 MoveCur(False) ;
 {$ENDIF}
{$IFNDEF SANSCOMPTA}
 ChargeDescriCPTA ;
{$ELSE SANSCOMPTA}
 {$IFDEF GCGC}
 ChargeDescriCPTA ;
 {$ENDIF GCGC}
{$ENDIF SANSCOMPTA}

 {=============== Exercices =====================}
 if Not ChargeMagExo(true) then PbExo:=True ;
 MoveCur(FALSE) ; ChargeTableLibre ;
 {$IFNDEF PGIIMMO}
 MoveCur(FALSE) ; ChargeHalleyUser ;
 MoveCur(FALSE) ; ChargeSousPlanAxe ;
 MoveCur(FALSE) ; ChargeCatBud ;
 ChargeparamsImmo ;
 ChargeJalEffet ;
 ChargeRoleMetier;
 {$ENDIF}
 CommitTrans ;
Except
 RollBack ;
 Result:=False ;
{$IFNDEF EAGLSERVER}
 PbBoot:=True ;
{$ENDIF EAGLSERVER}
End ;
DefStatus ;
FiniMove ;
{$IFNDEF PGIIMMO}
{$IFNDEF EAGLSERVER}
if PbBoot then HShowMessage('0;ATTENTION ! Problème à la connexion;Le chargement initial ne s''est pas correctement effectué;E;O;O;O;','','') ;
{$ENDIF}
if Pos('YYYY',uppercase(ShortDateFormat))<=0 then
   BEGIN
   StY:='ATTENTION ! Le format de dates est paramétré avec l''année sur 2 caractères !;Vous devez le modifier dans les paramètres régionaux de Windows du panneau de configuration' ;
   {$IFNDEF EAGLSERVER}
   HShowMessage('0;'+stY+';E;O;O;O;','','') ;
   {$ENDIF}
   END ;
{$ENDIF}

{$IFNDEF PGIIMMO}
{ CA - 25/08/2004 - Pour forcer le mode silencieux en PCL si un plan n'a pas déjà été déversé }
if not VH^.ModeSilencieux then
  VH^.ModeSilencieux := ((ctxPCL in V_PGI.PGIContexte) and (GetParamSocSecur('SO_NUMPLANREF',0)<=0));
if Not VH^.ModeSilencieux then
   BEGIN
//   if PbExo And ((CtxPcl in V_PGI.PGIContexte)=FALSE) then
   if PbExo And ((CtxPcl in V_PGI.PGIContexte)=FALSE) and (not (ctxLanceur in V_PGI.PGIContexte)) then
      //RRO if (ctxCompta in V_PGI.PGIContexte) or (ctxGescom in V_PGI.PGIContexte) then

      //Modif SBU - pas d'avertissement concernant l'exercice en CHR  si compta différée)
      if ((ctxCompta in V_PGI.PGIContexte) or (ctxGescom in V_PGI.PGIContexte))
          {$IFDEF CHR} and (not GetParamSocSecur('SO_GCDESACTIVECOMPTA',False)) {$ENDIF CHR} then
      BEGIN
      {$IFNDEF EAGLSERVER}
        {$IFDEF COMPTA}
        {$IFNDEF IMP}
        // FQ 13355 SBO 08/08/2005
        if ( GenDatesExoN(ExoEnCours) ) and ( TestExistExoNPlus1(ExoEnCours) <> 'OUV' ) then
          begin
          if PGIAsk('ATTENTION ! La date d''entrée ne correspond pas à un exercice ouvert (en Comptabilité). Vous ne pourrez pas créer de pièces, ni faire de traitements comptables. Voulez-vous ouvrir l''exercice ?', TitreHalley ) = mrYes then
            OuvertureExoModal;
          end
        else
        {$ENDIF}
        {$ENDIF COMPTA}
        { GC_DBR_GC15691 }
        HShowMessage('0;ATTENTION ! La date d''entrée ne correspond pas à un exercice ouvert; Vous ne pourrez pas faire de traitements comptables. Vérifiez la définition des exercices dans CEGID Comptabilité ou à défaut dans CEGID Administration sociétés;E;O;O;O;','','') ;
      {$ENDIF}
      END ;
   {$IFNDEF EAGLSERVER}
   if OkPb and (Not V_PGI.MajStructAuto) then HShowMessage('0;ATTENTION ! Paramétrage société incomplet;Vérifiez les caractères de bourrage dans les paramètres société;E;O;O;O;','','') ;
   If (Not ParamEuroOk) and (Not V_PGI.MajStructAuto) Then HShowMessage('0;ATTENTION ! Paramétrage société incomplet;Vérifiez les paramètres EURO dans la fiche société;E;O;O;O;','','') ;
   {$ENDIF}
   END ;
{$ENDIF}
// ajout me
(*
Q:=OpenSQL('SELECT DOS_CPPOINTAGESX,DOS_CPLIENGAMME,DOS_CPRDDATERECEP FROM DOSSIER where DOS_NODOSSIER="'+V_PGI.NoDossier+'"',TRUE) ;
if not Q.EOF then
begin
     VH^.CPPOINTAGESX := Q.FindField ('DOS_CPPOINTAGESX').asstring;
     VH^.CPLienGamme  := Q.FindField ('DOS_CPLIENGAMME').asstring;
     VH^.CPDaterecep  := Q.FindField ('DOS_CPRDDATERECEP').asdatetime;
end;
Ferme(Q) ;
*)
{$IFNDEF IMPEXP}
     VH^.CPPOINTAGESX := GetParamsoc ('SO_CPPOINTAGESX');
     VH^.CPLienGamme  := GetParamsoc ('SO_CPLIENGAMME');
     VH^.CPDaterecep  := GetParamsoc ('SO_CPRDDATERECEPTION');
{$ENDIF}
{$IFDEF COMPTA}
{JP 02/10/06 : Chargement du NoDossier du dossier Tréso}
ChargeDossierTreso;
{$ENDIF COMPTA}

{$IFDEF COMPTA} // GCO - 20/04/2007 - Révision intégrée Compta
  VH^.Revision.DossierPretSupervise := False;
  VH^.Revision.DossierSupervise     := False;
  VH^.Revision.Plan                 := GetParamSocSecur('SO_CPPLANREVISION', '', True);

  Q := OpenSQL('SELECT CIR_SUPERVISE, CIR_PRETSUPERVI FROM CREVINFODOSSIER WHERE ' +
               'CIR_NODOSSIER = "' + V_Pgi.NoDossier + '" AND ' +
               'CIR_EXERCICE = "' + VH^.EnCours.Code + '"', True ,-1,'',true);
  if not Q.Eof then
  begin
    VH^.Revision.DossierPretSupervise := (Q.FindField('CIR_PRETSUPERVI').AsString = 'X');
    VH^.Revision.DossierSupervise     := (Q.FindField('CIR_SUPERVISE').AsString = 'X');
  end;
  Ferme(Q);
{$ENDIF}
END ;
{$ENDIF}

{$IFNDEF NOVH}
{JP 02/10/06 : Chargement du NoDossier du dossier Tréso
{---------------------------------------------------------------------------------------}
procedure ChargeDossierTreso;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  T : Boolean;
begin
  T := (GetParamSocSecur('SO_TRBASETRESO', '') <> V_PGI.SchemaName) and (GetParamSocSecur('SO_TRBASETRESO', '') <> '');
  {Si la Tréso est multi sociétés et que l'on n'est pas sur la base de Treso ...}
  if EstComptaTreso and EstMultiSoc and T then begin
    {... On recherche le NoDossier de la base Tréso}
    Q := OpenSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_NOMBASE = "' + GetParamSocSecur('SO_TRBASETRESO', '') + '"', True,-1,'',true);
    if not Q.EOF then
      VH^.TresoNoDossier := Q.FindField('DOS_NODOSSIER').AsString;
    Ferme(Q);
  end
  else
    VH^.TresoNoDossier := V_PGI.NoDossier;
end;

PROCEDURE CHARGETVATPF ;
Var Q : TQuery ;
    NbBase : Integer ;
    lStDossier : string ;
BEGIN
VH^.TpfApres:=FALSE ; VH^.TpfDehors:=TRUE ;
{$IFDEF SPEC302}
FillChar(VH^.TabTVA,Sizeof(VH^.TabTva),#0) ; FillChar(VH^.TabTPF,Sizeof(VH^.TabTpf),#0) ;
Q:=OpenSQL('SELECT * FROM TXCPTTVA',TRUE,-1,'',true);
NbTVA:=0 ; NbTPF:=0 ;
While Not Q.EOF do
   BEGIN
   if Q.FindField('TV_TVAOUTPF').AsString=VH^.DefCatTVA then
      BEGIN
      Inc(NbTVA) ;
      VH^.TabTva[NbTva].TauxACH:=Q.FindField('TV_TAUXACH').AsFloat/100.0 ;
      VH^.TabTva[NbTva].TauxVTE:=Q.FindField('TV_TAUXVTE').AsFloat/100.0 ;
      VH^.TabTva[NbTva].CpteACH:=Q.FindField('TV_CPTEACH').AsString ;
      VH^.TabTva[NbTva].CpteVTE:=Q.FindField('TV_CPTEVTE').AsString ;
      VH^.TabTva[NbTva].EncaisACH:=Q.FindField('TV_ENCAISACH').AsString ;
      VH^.TabTva[NbTva].EncaisVTE:=Q.FindField('TV_ENCAISVTE').AsString ;
      VH^.TabTva[NbTva].Code:=Q.FindField('TV_CODETAUX').AsString ;
      VH^.TabTva[NbTva].Regime:=Q.FindField('TV_REGIME').AsString ;
      END else
      BEGIN
      Inc(NbTPF) ;
      VH^.TabTPF[NbTPF].TauxACH:=Q.FindField('TV_TAUXACH').AsFloat/100.0 ;
      VH^.TabTPF[NbTPF].TauxVTE:=Q.FindField('TV_TAUXVTE').AsFloat/100.0 ;
      VH^.TabTPF[NbTPF].CpteACH:=Q.FindField('TV_CPTEACH').AsString ;
      VH^.TabTPF[NbTPF].CpteVTE:=Q.FindField('TV_CPTEVTE').AsString ;
      VH^.TabTPF[NbTPF].EncaisACH:=Q.FindField('TV_ENCAISACH').AsString ;
      VH^.TabTPF[NbTPF].EncaisVTE:=Q.FindField('TV_ENCAISVTE').AsString ;
      VH^.TabTPF[NbTPF].Code:=Q.FindField('TV_CODETAUX').AsString ;
      VH^.TabTPF[NbTPF].Regime:=Q.FindField('TV_REGIME').AsString ;
      END ;
   Q.Next ;
   END ;
Ferme(Q) ;
{$ELSE}
  if VH^.LaTOBTVA<>Nil then
  BEGIN
    VH^.LaTOBTVA.Free ;
    VH^.LaTOBTVA:=TOB.Create('',Nil,-1) ;
  END ;

{$IFDEF EAGLCLIENT}
  AvertirCacheServer('TXCPTTVA'); // GCO - 11/09/2007 - FQ
{$ENDIF}
	Q:=OpenSQL('SELECT * FROM TXCPTTVA',TRUE,-1,'',true);
  if not Q.eof then VH^.LaTOBTVA.LoadDetailDB ('TXCPTTVA','','',Q,false,true);
  ferme (Q);
{$ENDIF}
{#TVAENC}
FillChar(VH^.NumCodeBase,Sizeof(VH^.NumCodeBase),#0) ;
// Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then lStDossier := TableToBase( 'NATCPTE' )
    else lStDossier := '' ;
Q:=OpenSelect('Select CC_CODE, CC_LIBRE from CHOIXCOD Where CC_TYPE="'+VH^.DefCatTVA+'" AND CC_LIBRE>"0" AND CC_LIBRE<="4"',lStDossier) ;
While Not Q.EOF do
   BEGIN
   NbBase:=StrToInt(Q.Fields[1].AsString) ;
   VH^.NumCodeBase[NbBase]:=Q.Fields[0].AsString ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;
{$ENDIF}


{$IFNDEF NOVH}
FUNCTION TVA2CPTE(ModeTVA,Tva : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTVA,ModeTVA,TVA],False) ;
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEACH')
        else
        Result:=TOBT.GetValue('TV_CPTACHFARFAE');
    end else
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEVTE')
        else
        Result:=TOBT.GetValue('TV_CPTVTEFARFAE');
    end;
  END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TVA2ENCAIS(ModeTVA,Tva : String3 ; Achat : Boolean; RG : boolean=false; FarFae : boolean=false) : String ;
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTVA,ModeTVA,TVA],False) ;
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not RG then
      begin
        if not FarFae then
          Result := TOBT.GetValue('TV_ENCAISACH')
          else
          Result := TOBT.GetValue('TV_CPTACHFARFAE')
      end else
        Result:=TOBT.GetValue('TV_CPTACHRG');
    end else
    begin
      if not RG then
      begin
        if not FarFae then
          Result:=TOBT.GetValue('TV_ENCAISVTE')
          else
          Result := TOBT.GetValue('TV_CPTVTEFARFAE')
      end else
        Result:=TOBT.GetValue('TV_CPTVTERG')
    end;
  END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TVA2TAUX(ModeTVA,Tva : String3 ; Achat : Boolean) : Real ;
Var TOBT : TOB ;
BEGIN
Result:=0 ;
TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTVA,ModeTVA,TVA],False) ;
if TOBT<>Nil then
    BEGIN
    if Achat then Result:=TOBT.GetValue('TV_TAUXACH')/100.0
             else Result:=TOBT.GetValue('TV_TAUXVTE')/100.0 ;
    END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TPF2CPTE(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTPF,ModeTVA,TPF],False) ;
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEACH')
        else
        Result := TOBT.GetValue('TV_CPTACHFARFAE');
    end else
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEVTE')
        else
        Result := TOBT.GetValue('TV_CPTVTEFARFAE');
    end;
  END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TPF2ENCAIS(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTPF,ModeTVA,TPF],False) ;
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_ENCAISACH')
        else
        Result := TOBT.GetValue('TV_CPTACHFARFAE');
    end else
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_ENCAISVTE')
        else
        Result := TOBT.GetValue('TV_CPTVTEFARFAE');
    end;
  END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TPF2TAUX(ModeTVA,Tpf : String3 ; Achat : Boolean) : Real ;
Var TOBT : TOB ;
BEGIN
Result:=0 ;
TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTPF,ModeTVA,TPF],False) ;
if TOBT<>Nil then
    BEGIN
    if Achat then Result:=TOBT.GetValue('TV_TAUXACH')/100.0
             else Result:=TOBT.GetValue('TV_TAUXVTE')/100.0 ;
    END ;
END ;
{$ENDIF}

{============================================================================}
FUNCTION  TOTDIFFERENT(X1,X2 : Double) : BOOLEAN ;
BEGIN
Result:=Not (Abs(X1-X2)<0.000001) ;
END ;

{============================================================================}
{$IFNDEF NOVH}
FUNCTION HT2TVA ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
Var TpfFRS : Double ;
BEGIN
TpfFrs:=0 ;
if (VH^.PaysLocalisation=CodeISOES) or (Not SoumisTPF) then //XMG 24/02/2005
   BEGIN
   HT2TVA:=Arrondi(THT*Tva2Taux(ModeTVA,Tva,Achat),Dec) ;
   END else
   BEGIN
   if VH^.TpfDehors then
      BEGIN
      if VH^.TpfApres then HT2TVA:=Arrondi(THT*Tva2Taux(ModeTVA,Tva,Achat),Dec)
                     else HT2TVA:=Arrondi(THT*((1.0+Tpf2Taux(ModeTVA,Tpf,Achat))+TpfFRS)*Tva2Taux(ModeTVA,Tva,Achat),Dec) ;
      END else
      BEGIN
      HT2TVA:=Arrondi(THT*(((1.0+Tpf2Taux(ModeTVA,Tpf,Achat))*Tpf2Taux(ModeTVA,Tpf,Achat))+1)*Tva2Taux(ModeTVA,Tva,Achat),Dec) ;
      END ;
   END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION HT2TPF ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
Var TpfFRS : Double ;
BEGIN
TpfFrs:=0 ; HT2Tpf:=0 ; if Not SoumisTpf then Exit ;
if VH^.PaysLocalisation=CodeISOES then //XMG 24/02/2005
   HT2TPF:=Arrondi(THT*(Tpf2Taux(ModeTVA,Tpf,Achat)),Dec)
else
if VH^.TpfDehors then
   BEGIN
   if VH^.TpfApres then HT2TPF:=Arrondi(THT*(1.0+Tva2Taux(ModeTVA,Tva,Achat))*Tpf2Taux(ModeTVA,Tpf,Achat)+TpfFRS,Dec)
                   else HT2TPF:=Arrondi(THT*Tpf2Taux(ModeTVA,Tpf,Achat)+TPFFRS,Dec) ;
   END else
   BEGIN
   HT2TPF:=Arrondi(THT*(1.0+Tpf2Taux(ModeTVA,Tpf,Achat))*Tpf2Taux(ModeTVA,Tpf,Achat),Dec) ;
   END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION HT2TTC ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
Var HTTPF : real ;
    TpfFRS : Double ;
BEGIN
TpfFrs:=0 ;
if Not SoumisTPF then
   BEGIN
   HT2TTC:=Arrondi(THT*(1.0+Tva2Taux(ModeTva,Tva,Achat)),Dec) ;
   END else
   BEGIN
   if VH^.PaysLocalisation=CodeISOES then //XMG 24/02/2005
      HT2TTC:=Arrondi(THT*(1.0+(Tva2Taux(ModeTva,Tva,Achat)+Tpf2Taux(ModeTVA,Tpf,Achat))),Dec)
   else
   if VH^.TpfDehors then
      BEGIN
      if VH^.TpfApres then
         BEGIN
         HT2TTC:=Arrondi(THT*(1.0+Tva2Taux(ModeTVA,Tva,Achat))*(1.0+Tpf2Taux(ModeTVA,Tpf,Achat))+TpfFRS,Dec) ;
         END else
         BEGIN
         HTTPF:=THT*Tpf2Taux(ModeTVA,Tpf,Achat)+TpfFRS ;
         HT2TTC:=Arrondi((THT+HTTPF)*(1.0+Tva2Taux(ModeTVA,Tva,Achat)),Dec) ;
         END ;
      END else
      BEGIN
      HTTPF:=THT*(1.0+Tpf2Taux(ModeTVA,Tpf,Achat))*Tpf2Taux(ModeTVA,Tpf,Achat) ;
      HT2TTC:=Arrondi((THT+HTTPF)*(1.0+Tva2Taux(ModeTVA,Tva,Achat)),Dec) ;
      END ;
   END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TTC2HT ( TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
Var IntFR,CoefTpf : real ;
    TpfFRS : Double ;
BEGIN
TpfFrs:=0 ;
if Not SoumisTPF then
   BEGIN
   TTC2HT:=Arrondi(TTTC/(1.0+Tva2Taux(ModeTva,Tva,Achat)),Dec) ;
   END else
   BEGIN
   if VH^.PaysLocalisation=CodeISOES then //XMG 24/02/2005
      TTC2HT:=Arrondi(TTTC/(1.0+(Tva2Taux(ModeTva,Tva,Achat)+Tpf2Taux(ModeTVA,Tpf,Achat))),Dec)
   else
   if VH^.TpfDehors then
      BEGIN
      if VH^.TpfApres then
         BEGIN
         TTC2HT:=Arrondi((TTTC-TPFFRS)/(1.0+Tva2Taux(ModeTVA,Tva,Achat))/(1.0+Tpf2Taux(ModeTVA,Tpf,Achat)),Dec) ;
         END else
         BEGIN
         INTFR:=TPFFRS*(1.0+Tva2Taux(ModeTVA,Tva,Achat)) ;
         TTC2HT:=Arrondi((TTTC-IntFR)/(1.0+Tva2Taux(ModeTVA,Tva,Achat))/(1.0+Tpf2Taux(ModeTVA,Tpf,Achat)),Dec) ;
         END ;
      END else
      BEGIN
      CoefTPF:=1.0+(Tpf2Taux(ModeTVA,Tpf,Achat)*(1.0+Tpf2Taux(ModeTVA,Tpf,Achat))) ;
      TTC2HT:=Arrondi(TTTC/(1.0+Tva2Taux(ModeTVA,Tva,Achat))/CoefTPF,Dec) ;
      END ;
   END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TTC2TPF ( TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
BEGIN
TTC2TPF:=HT2TPF(TTC2HT(TTTC,ModeTVA,SoumisTPF,Tva,Tpf,Achat,Dec),modeTVA,SoumisTPF,Tva,Tpf,Achat,Dec) ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
FUNCTION TTC2TVA (TTTC : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
BEGIN
TTC2TVA:=HT2TVA(TTC2HT(TTTC,ModeTVA,SoumisTPF,Tva,Tpf,Achat,Dec),modeTVA,SoumisTPF,Tva,Tpf,Achat,Dec) ;
END ;
{$ENDIF}


{$IFDEF EAGLCLIENT}
{$ELSE}
Procedure ChgMaskChamp (C : TFloatField ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) ;
Var j      : integer ;
BEGIN
if Decim=0 then C.DisplayFormat:='#,##0' else C.DisplayFormat:='#,##0.' ;
for j:=1 to Decim do C.DisplayFormat:=C.DisplayFormat+'0' ;
if AffSymb then
   BEGIN
   if IsSolde
      then BEGIN C.DisplayFormat:=C.DisplayFormat+' '+Symbole+';-'+C.DisplayFormat+' '+Symbole+';'+C.DisplayFormat+' '+Symbole ; END
      else BEGIN C.DisplayFormat:=C.DisplayFormat+' '+Symbole+';-'+C.DisplayFormat+' '+Symbole+';''''' ; END;
   END else
   BEGIN
   if IsSolde
      then BEGIN C.DisplayFormat:=C.DisplayFormat+';-'+C.DisplayFormat+';'+C.DisplayFormat ; END
      else BEGIN C.DisplayFormat:=C.DisplayFormat+';-'+C.DisplayFormat+';''''' ; END;
   END ;
END ;
{$ENDIF}


Function MontantToStr (Montant : Double ; Decim : Integer ; AffSymb : boolean ; Symbole : String ; IsSolde : Boolean) : string ;
Var Chaine : string ;
BEGIN
if AffSymb then
   BEGIN
   if IsSolde then
      BEGIN
      Chaine:=FloatToStrF(Montant,ffNumber,20,Decim)+' '+Symbole ;
      END else
      BEGIN
      if Montant=0 then Chaine:=''
                   else Chaine:=FloatToStrF(Montant,ffNumber,20,Decim)+' '+Symbole ;
      END ;
   END else
   BEGIN
   if IsSolde then
      BEGIN
      Chaine:=FloatToStrF(Montant,ffNumber,20,Decim) ;
      END else
      BEGIN
      if Montant=0 then Chaine:=''
                   else Chaine:=FloatToStrF(Montant,ffNumber,20,Decim) ;
      END ;
   END ;
MontantToStr:=Chaine ;
END ;


Procedure CorrespToCombo(Var CB : THValComboBox ; Fb : TFichierBase );
Var Q1 : TQuery ;
    StAnd : String ;
BEGIN
StAnd:='' ;
Case Fb of
    fbGene               :  StAnd:='and (CO_CODE="GE1" or CO_CODE="GE2")' ;
    fbAux                :  StAnd:='and (CO_CODE="AU1" or CO_CODE="AU2")' ;
    fbAxe1..fbAxe5       :  StAnd:='and CO_CODE like "'+FbToAxe(Fb)+'%" ' ;
    End ;
CB.Items.Clear ; CB.Values.Clear ;
Q1:=OpenSQL('Select CO_CODE, CO_LIBELLE from COMMUN where CO_TYPE="CTC" '+StAnd+' order by CO_CODE ',True,-1,'',true);
While Not Q1.EOF do
    BEGIN
    CB.Items.Add(Q1.FindField('CO_LIBELLE').AsString) ;
    CB.Values.Add(Q1.FindField('CO_CODE').AsString) ;
    Q1.Next ;
    END ;
Ferme(Q1) ;
END ;

Procedure CorrespToCodes(Plan : THValComboBox ; Var C1, C2 : TComboBox);
Var St : String ;
    Q1 : TQuery ;
BEGIN
C1.Clear ; C2.Clear ;
St:='Select CR_CORRESP, CR_LIBELLE From CORRESP Where CR_TYPE="'+Plan.Value+'" '
   +'Order By CR_TYPE, CR_CORRESP' ;
Q1:=OpenSQL(St,True,-1,'',true); ; Q1.First ;
While Not Q1.EOF do
    BEGIN     {A voir si on ajoute les Codes en value et les libellés en Items ... }
    C1.Items.Add(Q1.Fields[0].AsString) ;
    C2.Items.Add(Q1.Fields[0].AsString) ;
    Q1.Next ;
    END ;
Ferme(Q1) ;
END ;

Procedure RuptureToCodes(Plan : THValComboBox ; Var C1, C2 : TComboBox ; Fb : TFichierBase );
Var St, StAnd, StAxe : String ;
    Q1 : TQuery ;
BEGIN
C1.Clear ; C2.Clear ;
StAnd:='' ;
Case Fb of
    fbGene               :  StAnd:='RU_NATURERUPT="RUG"' ;
    fbAux                :  StAnd:='RU_NATURERUPT="RUT"' ;
    fbAxe1..fbAxe5       :  BEGIN StAxe:=FbToAxe(Fb) ; StAnd:='RU_NATURERUPT="RU'+Char(StAxe[2])+'" ' ; END ;
    End ;
St:='Select RU_CLASSE, RU_LIBELLECLASSE From Rupture Where '+StAnd+' and RU_PLANRUPT="'+Plan.Value+'" '
   +'Order By RU_CLASSE' ;
Q1:=OpenSQL(St,True,-1,'',true); ; Q1.First ;
While Not Q1.EOF do
    BEGIN     {A voir si on ajoute les Codes en value et les libellés en Items ... }
    C1.Items.Add(Q1.Fields[0].AsString) ;
    C2.Items.Add(Q1.Fields[0].AsString) ;
    Q1.Next ;
    END ;
Ferme(Q1) ;
END ;

Procedure CodesRuptSec(Var LCodes : TStringlist ; Plan,Axe : String) ;
Var Q1 : TQuery ;
BEGIN
Q1:=OpenSQL('Select RU_CLASSE from Rupture where RU_NATURERUPT="RU'+Char(Axe[2])+'" and RU_PLANRUPT="'+Plan+'" '
          +'Order by RU_CLASSE ',TRUE,-1,'',true);
Q1.First ;
if LCodes<>Nil then LCodes.Clear ;
While Not Q1.Eof Do
    BEGIN
    LCodes.Add(Q1.Fields[0].AsString) ;
    Q1.Next ;
    END ;
Ferme(Q1) ;
END ;

Procedure RetoucheSolde(AffSymb : Boolean ; DC : Char ; Symbole : String ; Var LeResultat : String) ;
BEGIN
if AffSymb then LeResultat:=LeResultat+' '+Symbole+' '+DC else LeResultat:=LeResultat+' '+DC ;
END ;

Function PrintSolde(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean) : String ;
var LaValeur : Double ;
    LeResultat : String ;
    OnInverse : Boolean ;
BEGIN
LaValeur:=0 ; OnInverse:=FALSE ;
if (TD=TC) then
   BEGIN
   LaValeur:=TD-TC ;
   LeResultat:=FloatToStrF(LaValeur,ffNumber,20,Decim) ;
   if AffSymb then PrintSolde:=LeResultat+' '+Symbole+' '+' ' else PrintSolde:=LeResultat+'  '+' ' ;
   END else
if (Abs(TD)>=Abs(TC)) then
   BEGIN
   LaValeur:=TD-TC ; If LaValeur<0 Then OnInverse:=TRUE ;
   LeResultat:=FloatToStrF(Abs(LaValeur),ffNumber,20,Decim) ;
   If OnInverse Then RetoucheSolde(AffSymb,'C',Symbole,LeResultat) Else RetoucheSolde(AffSymb,'D',Symbole,LeResultat) ;
//   if AffSymb then PrintSolde:=LeResultat+' '+Symbole+' '+'D' else PrintSolde:=LeResultat+' '+'D' ;
   END else
if Abs(TD)<Abs(TC) then
   BEGIN
   LaValeur:=TC-TD ; If LaValeur<0 Then OnInverse:=TRUE ;
   LeResultat:=FloatToStrF(Abs(LaValeur),ffNumber,20,Decim) ;
   If OnInverse Then RetoucheSolde(AffSymb,'D',Symbole,LeResultat) Else RetoucheSolde(AffSymb,'C',Symbole,LeResultat) ;
//   while Pos('-', LeResultat) > 0 do LeResultat[Pos('-', LeResultat)] := ' ';
//   if AffSymb then PrintSolde:=LeResultat+' '+Symbole+' '+'C' else PrintSolde:=LeResultat+' '+'C' ;
   END ;
If LaValeur=0 then PrintSolde:='' Else PrintSolde:=LeResultat ;
END ;

Function PrintEcart(TD,TC : DOUBLE ; Decim : Integer ; DebitPos : Boolean) : String ;
var LaValeur   : Double ;
BEGIN
LaValeur:=0 ;
if (TD=TC) then
   BEGIN
   LaValeur:=Arrondi(TD-TC,Decim) ; Result:=FloatToStrF(LaValeur,ffNumber,20,Decim)+'   ' ;
   END else
if Abs(TD)>=Abs(TC) then
   BEGIN
   LaValeur:=Arrondi(TD-TC,Decim) ; If Not DebitPos Then LaValeur:=-1*LaValeur ;
   Result:=FloatToStrF(LaValeur,ffNumber,20,Decim)+'   ' ;
   END else
if Abs(TD)<Abs(TC) then
   BEGIN
   LaValeur:=Arrondi(TC-TD,Decim) ; If DebitPos Then LaValeur:=-1*LaValeur ;
   Result:=FloatToStrF(LaValeur,ffNumber,20,Decim)+'   ' ;
//   while Pos('-', LeResultat) > 0 do LeResultat[Pos('-', LeResultat)] := ' ';
   END ;
If LaValeur=0 then Result:='' ;
END ;

Function PrintSoldeFormate(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean ; FormMont : String) : String ;
var LaValeur   : Double ;
    LeResultat, StSym : String ;
    NegToPos : Boolean ;
BEGIN
LaValeur:=0 ; NegToPos:=FALSE ;
if (TD=TC) then
   BEGIN
   LaValeur:=TD-TC ;
   LeResultat:=FloatToStrF(LaValeur,ffNumber,20,Decim) ;
   if AffSymb then PrintSoldeFormate:=LeResultat+' '+Symbole+' '+' ' Else PrintSoldeFormate:=LeResultat+'  '+' '
   END else
if Abs(TD)>=Abs(TC) then
   BEGIN
   LaValeur:=TD-TC ; If (LaValeur<0) And (FormMont='PC') Then BEGIN LaValeur:=LaValeur*(-1) ; NegToPos:=TRUE ; END ;
   LeResultat:=FloatToStrF(LaValeur,ffNumber,20,Decim) ;
   If AffSymb then StSym:=' '+Symbole else StSym:=' ' ;
   PrintSoldeFormate:=LeResultat+StSym+' '+'D' ;
   If FormMont='PC' then
     BEGIN { Crédit positif }
     If NegToPos Then PrintSoldeFormate:=' '+LeResultat+StSym+' '+' '
                 Else PrintSoldeFormate:='-'+LeResultat+StSym+' '+' ' ;
     END else
     If FormMont='PD' then PrintSoldeFormate:=LeResultat+StSym+' '+' ' Else { Débit positif }
       If FormMont='DP' then PrintSoldeFormate:='('+LeResultat+')'+StSym+' '+' ' Else{ Debit parenthèse }
          If FormMont='CP' then PrintSoldeFormate:=LeResultat+StSym+' '+' ' ;{ Crédit parenthèse }
   END else
if Abs(TD)<Abs(TC) then
   BEGIN
   LaValeur:=TC-TD ; If (LaValeur<0) And (FormMont='PD') Then BEGIN LaValeur:=LaValeur*(-1) ; NegToPos:=TRUE ; END ;
   LeResultat:=FloatToStrF(LaValeur,ffNumber,20,Decim) ;
//   while Pos('-', LeResultat) > 0 do LeResultat[Pos('-', LeResultat)] := ' ';
   If AffSymb then StSym:=' '+Symbole else StSym:=' ' ;
   PrintSoldeFormate:=LeResultat+StSym+' '+'C' ;
   If FormMont='PD' then
     BEGIN { Débit positif }
     If NegToPos Then PrintSoldeFormate:=' '+LeResultat+StSym+' '+' '
                 Else PrintSoldeFormate:='-'+LeResultat+StSym+' '+' ' ;
     END Else
     If FormMont='CP' then PrintSoldeFormate:='('+LeResultat+')'+StSym+' '+' ' Else{ Crédit parenthèse }
       If FormMont='PC' then PrintSoldeFormate:=LeResultat+StSym+' '+' ' else { Crédit positif }
          If FormMont='DP' then PrintSoldeFormate:=LeResultat+StSym+' '+' ' Else{ Debit parenthèse }
   END ;
If LaValeur=0 then PrintSoldeFormate:='' ;

END ;

Function  PrintSolde2(TD,TC : DOUBLE ; Decim : Integer ; Symbole : String ; AffSymb : boolean ; FormMont : THValCombobox) : String ;
var LaValeur : Double ;
    LeResultat, ForMon : String ;
    Coef     : Double ;
BEGIN
PrintSolde2:='' ; Coef:=1.0 ;
If Not AffSymb then Symbole:='' ;
If FormMont.Value='N' then begin ForMon:='' ; Coef:=1.0 ; end else
If FormMont.Value='K' then begin ForMon:='' ; Coef:=1000.0 ; end else
If FormMont.Value='M' then begin ForMon:='' ; Coef:=1000000.0 ; end ;
if (TD=TC) then
   BEGIN
   LaValeur:=((TD-TC)/Coef) ;
   LeResultat:=FloatToStrF(LaValeur,ffNumber,20,Decim) ;
   If LaValeur<>0 then PrintSolde2:=LeResultat+' '+ForMon+Symbole+' '+' ' ;
   END else if Abs(TD)>=Abs(TC) then
   BEGIN
   LaValeur:=(TD-TC)/Coef ;
   LeResultat:=FloatToStrF(LaValeur,ffNumber,20,Decim) ;
   If LaValeur<>0 then PrintSolde2:=LeResultat+' '+ForMon+Symbole+' '+'D' ;
   END else if Abs(TD)<Abs(TC) then
   BEGIN
   LaValeur:=(TC-TD)/Coef ;
   LeResultat:=FloatToStrF(LaValeur,ffNumber,20,Decim) ;
//   while Pos('-', LeResultat) > 0 do LeResultat[Pos('-', LeResultat)] := ' ';
   If LaValeur<>0 then PrintSolde2:=LeResultat+' '+ForMon+Symbole+' '+'C' ;
   END ;
END ;

function MoinsSymbole(LeSolde : String) : double ;
Var i : Integer ;
    Alors : String ;                             { Déformate les montants }
BEGIN
Alors:='' ;
if LeSolde='' then LeSolde:='0' ;
for i:=1 to Length(LeSolde) do
    BEGIN
    if LeSolde[i] in ['0'..'9','.',',','-'] then Alors:=Alors+LeSolde[i] ;
    END ;
MoinsSymbole:=StrToFloat(Alors) ;
END ;

Function  SqlCptInterdit(LeChamp : String ; Var PourSql : String ; ZoneInterdit : TEdit) : Boolean ;
const Ensemble=['0'..'9','A'..'Z',';',':'] ;
      Separateur=';' ;
var   i,J,T      : integer ;
      Cpte,St,Machin,UnCpt,LaSql : string ;
      Okok, Trouve    : boolean ;
      //BourreCpt       : Integer ;
begin
PourSql:='' ;
if ZoneInterdit.Text='' then  BEGIN Result:=True ; Exit ; END ;
(* Probleme pour Section ---> Axe !
If LeChamp='G_GENERAL' then BourreCpt:=V_PGI.Cpta[fbGene].lg else
If LeChamp='T_AUXILIAIRE' then BourreCpt:=V_PGI.Cpta[fbAux].lg else

If LeChamp='S_SECTION' then BourreCpt:=V_PGI.Cpta[AxeToFb(Cp.Axe)].lg ;
*)
Cpte:=ZoneInterdit.Text ;
if Cpte[Length(Cpte)]<>Separateur then BEGIN Cpte:=Cpte+';' ;  ZoneInterdit.Text:=Cpte ; END ;
Okok:=True ; St:=ReadTokenSt(Cpte) ;
While ((St<>'') and (Okok)) do
   BEGIN
   for i:=1 to Length(St)-1 do if not (St[i] in Ensemble) then Okok:=False ;
   if Okok then
      BEGIN
      Trouve:=False ;
      for J:=1 to Length(St) do if St[J] in [':'] then  BEGIN Trouve:=True ; Break ; END ;
      If Trouve Then
         BEGIN
         T:=1 ; UnCpt:=Fourchette(St) ;
         While ( (UnCpt<>'') and (T<3) ) do
               BEGIN
               if T=1 then Machin:=' And '+LeChamp+' Not BETWEEN "'+UnCpt+'"' Else Machin:=Machin+' And "'+UnCpt+'_'+'"' ;
               T:=T+1 ; UnCpt:=Fourchette(St) ;
               END ;
         LaSql:=Machin ;
         End Else Begin LaSql:=' And '+LeChamp+'<>"'+St+'"' end ;
      PourSql:=PourSql+LaSql ;
      St:=ReadTokenSt(Cpte) ;
      END ;

   END ;
Result:=OkOk ;
END ;

FUNCTION Fourchette( Var St : String) : String ;
Var i : Integer ;
BEGIN
i:=Pos(':',St) ; if i<=0 then i:=Length(St)+1 ; Result:=Copy(St,1,i-1) ; Delete(St,1,i) ;
END ;

{$IFNDEF NOVH}
{JP 14/08/07 : FQ 20091 : gestion du cas où Exo est vide. Ce n'est pas fait dans ExoToDates, car
               personne n'est trop sûr sur d'éventuels impacts
{---------------------------------------------------------------------------------------}
procedure ExoToEdDates(Exo : String3; ED1, ED2 : TControl);
{---------------------------------------------------------------------------------------}
begin
  if Exo = '' then begin
    TEdit(ED1).Text := StDate1900;
    TEdit(ED2).Text := StDate2099;
  end
  else
    ExoToDates(Exo, ED1, ED2);
end;

procedure ExoToDates ( Exo : String3 ; ED1,ED2 : TControl ) ;
Var D1,D2 : TDateTime ;
    Q     : TQuery;
    Okok  : boolean ;
BEGIN
{JP 14/08/07 : FQ 20091 : Si l'on veut gérer le cas où Exo vaut '', appeler ExoToEdDates définie ci-dessus}
if EXO='' then Exit ;
Okok:=True ; D1:=Date ; D2:=Date ;
If EXO=VH^.Precedent.Code Then BEGIN D1:=VH^.Precedent.Deb ; D2:=VH^.Precedent.Fin ; END Else
If EXO=VH^.EnCours.Code Then BEGIN D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ; END Else
If EXO=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END Else
   BEGIN
   Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE,-1,'',true);
   if Not Q.EOF then
      BEGIN
      D1:=Q.FindField('EX_DATEDEBUT').asDateTime ; D2:=Q.FindField('EX_DATEFIN').asDateTime ;
      END else Okok:=False ;
   Ferme(Q) ;
   END;
if Okok then BEGIN TEdit(ED1).Text:=DateToStr(D1) ; TEdit(ED2).Text:=DateToStr(D2) ; END ;
END ;
{$ENDIF} 

FUNCTION QUELEXODTBUD(DD : TDateTime) : String ;
Var i : Integer ;
    Q : Tquery ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelExoDtBUD') ;
{$ENDIF}
Result := GetEnCours.Code ;
If (dd>= GetEnCours.Deb) and (dd <= GetEnCours.Fin) then Result := GetEnCours.Code else
   If (dd >= GetSuivant.Deb) and (dd <= GetSuivant.Fin) then Result := GetSuivant.Code Else
      If (dd >= GetPrecedent.Deb) and (dd <= GetPrecedent.Fin) then Result := GetPrecedent.Code Else
      BEGIN
      For i:=1 To 5 Do
        BEGIN
        If (dd >= GetExoClo[i].Deb) And (dd <= GetExoClo[i].Fin)then BEGIN Result := GetExoClo[i].Code ; Exit ; END ;
        END ;
      Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_DATEDEBUT<="'+USDATETIME(DD)+'" AND EX_DATEFIN>="'+USDATETIME(DD)+'" ' ,TRUE,-1,'',true);
      if Not Q.EOF then Result:=Q.FindField('EX_EXERCICE').AsString ;
      Ferme(Q) ;
      END ;
end ;


Function SQLPremierDernier(fb : TFichierbase ; Prem : Boolean) : String ;
Var Q       : TQuery ;
    SQL,Cpt,LaTable,Desc,where : String ;
begin
result:='' ;
If Prem Then Desc:=' ASC ' Else Desc:=' DESC ' ;
Case fb Of
  fbGene : BEGIN
           Cpt:='G_GENERAL' ; LaTable:='GENERAUX' ; Where:=' ' ;
           END ;
  fbAux  : BEGIN
           Cpt:='T_AUXILIAIRE' ; LaTable:='TIERS' ; Where:=' ' ;
           END ;
  fbJAL  : BEGIN
           Cpt:='J_JOURNAL' ; LaTable:='JOURNAL' ; Where:=' ' ;
           END ;
  fbAxe1..fbAxe5 : BEGIN
                    if VH^.LiaisonY2ViaShare then
                    begin
                      Cpt:='CSP_SECTION' ; LaTable:='CSECTION' ; Where:=' WHERE CSP_AXE="'+'A'+IntToStr(ord(fb)+1)+'" ' ;
                    end else
                    begin
                      Cpt:='S_SECTION' ; LaTable:='SECTION' ; Where:=' WHERE S_AXE="'+'A'+IntToStr(ord(fb)+1)+'" ' ;
                    end;
                   END ;
  fbBudgen : BEGIN
             Cpt:='BG_BUDGENE' ; LaTable:='BUDGENE' ; Where:=' ' ;
             END ;
  fbBudSec1..fbBudSec5 : BEGIN
             Cpt:='BS_BUDSECT' ; LaTable:='BUDSECT' ; Where:=' WHERE BS_AXE="'+'A'+IntToStr(ord(fb)-13)+'" ' ;
             END ;
  fbBudJal : BEGIN
             Cpt:='BJ_BUDJAL' ; LaTable:='BUDJAL' ; Where:=' ' ;
             END ;
  END ;
SQL:='SELECT '+Cpt+' FROM '+LaTable+where+' ORDER BY '+Cpt+desc ;
Q:=OpenSQL(SQL,TRUE,-1,'',true);
if Not Q.Eof then
   BEGIN
   Case fb Of
     fbGene : Result:=Q.FindField('G_GENERAL').AsString ;
     fbAux  : Result:=Q.FindField('T_AUXILIAIRE').AsString ;
     fbAxe1..fbAxe5 : Result:=Q.FindField('S_SECTION').AsString ;
     fbJal  : Result:=Q.FindField('J_JOURNAL').AsString ;
     fbBudgen : Result:=Q.FindField('BG_BUDGENE').AsString ;
     fbBudSec1..fbBudSec5 : Result:=Q.FindField('BS_BUDSECT').AsString ;
     fbBudJal : Result:=Q.FindField('BJ_BUDJAL').AsString ;
     end ;
   END ;
Ferme(Q) ;
end ;


(*======================================================================*)
Procedure PremierDernier(fb : TFichierBase ; Var Cpt1,Cpt2 : String ) ;

begin
Cpt1:=SQLPremierDernier(fb,TRUE) ;
Cpt2:=SQLPremierDernier(fb,FALSE) ;
end ;

(*======================================================================*)
Function SQLPremierDernierRub(ZoomTable : TZoomTable ; SynPlus : String ; Prem : Boolean) : String ;
Var Q       : TQuery ;
    SQL,Cpt,LaTable,Desc,where : String ;
begin
result:='' ;
If Prem Then Desc:=' ASC ' Else Desc:=' DESC ' ;
Cpt:='RB_RUBRIQUE' ; LaTable:='RUBRIQUE' ; Where:='' ;
Case ZoomTable of
  tzRubCPTA  : Where:=' WHERE (RB_NATRUB="CPT")' ;
  tzRubBUDG  : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBG%" ' ;
  tzRubBUDS  : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBS%" ' ;
  tzRubBUDSG : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "S/G%" ' ;
  tzRubBUDGS : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "G/S%" ' ;
  Else Exit ;
  END ;
If (SynPlus<>'') Then
  BEGIN
  If ZoomTable=tzRubCPTA Then Where:=Where+'AND RB_FAMILLES like "%'+SynPlus+'%"'
                         Else Where:=Where+'AND RB_BUDJAL like "%'+SynPlus+'%"' ;
  END ;
SQL:='SELECT '+Cpt+' FROM '+LaTable+where+' ORDER BY '+Cpt+desc ;
Q:=OpenSQL(SQL,TRUE,-1,'',true);
if Not Q.Eof then Result:=Q.FindField('RB_RUBRIQUE').AsString ;
Ferme(Q) ;
end ;

(*======================================================================*)
Procedure PremierDernierRub(ZoomTable : TZoomTable ; SynPlus : String ; Var Cpt1,Cpt2 : String ) ;
begin
Cpt1:=SQLPremierDernierRub(ZoomTable,SynPlus,TRUE) ;
Cpt2:=SQLPremierDernierRub(ZoomTable,SynPlus,FALSE) ;
end ;

{$IFNDEF NOVH}
Function  JaiLeDroitNature ( Fichier,Niveau,Action : integer ; Parle : boolean ) : boolean ;
BEGIN
Result:=VH^.NivTableLibre[Fichier,Niveau][Action]<>'-' ;
{$IFNDEF EAGLSERVER}
if (Parle) And (Not Result) then HShowMessage('0;Action interdite !;Vous n''avez pas le droit d''effectuer cette opération;W;O;O;O','','') ;
{$ENDIF}
END ;
{$ENDIF}

Function QuelDateDeExo(CodeExo : String3 ; Var Exo : TExoDate) : Boolean ;
Var i : Integer ;
    Q : TQuery ;
BEGIN
{$IFDEF NOVH}
Exo:=GetEnCours ; Result:=FALSE ;
if CodeExo=GetPrecedent.Code then BEGIN Result:=TRUE ; Exo:=GetPrecedent ; END else
   if CodeExo=GetEnCours.Code then BEGIN Result:=TRUE ; Exo:=GetEnCours ; END else
      if CodeExo=GetSuivant.Code then BEGIN Result:=TRUE ; Exo:=GetSuivant ; END Else
         BEGIN
         For i:=1 To 5 Do BEGIN If CodeExo=GetExoClo[i].Code then
            BEGIN Exo:=GetExoClo[i] ; Result:=TRUE ; Exit ; END ; END ;
          // ajout me
              if TCPContexte.GetCurrent.Exercice.Exercices[1].Code <> '' then
              begin
                    for  i:= 1  to TCPContexte.GetCurrent.Exercice.NbExercices do
                    begin
                                if (TCPContexte.GetCurrent.Exercice.Exercices[i].Code = CodeExo) then
                                begin
                                     Exo.Code    := TCPContexte.GetCurrent.Exercice.Exercices[i].Code ;
                                     Exo.Deb     := TCPContexte.GetCurrent.Exercice.Exercices[i].Deb ;
                                     Exo.Fin     := TCPContexte.GetCurrent.Exercice.Exercices[i].Fin ;
                                     Result      := TRUE ;
                                     break;
                                end;
                    end;
              end
              else
              begin
                   Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="'+CodeExo+'"',TRUE,-1,'',true);
                   If Not Q.Eof Then
                      BEGIN
                      Exo.Code:=Q.FindField('EX_EXERCICE').AsString ;
                      Exo.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
                      Exo.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
                      Result:=TRUE ;
                      END ;
                   Ferme(Q) ;
              end;
        END ;
{$ELSE}
Exo:=VH^.EnCours ; Result:=FALSE ;
if CodeExo=VH^.Precedent.Code then BEGIN Result:=TRUE ; Exo:=VH^.Precedent ; END else
   if CodeExo=VH^.EnCours.Code then BEGIN Result:=TRUE ; Exo:=VH^.EnCours ; END else
      if CodeExo=VH^.Suivant.Code then BEGIN Result:=TRUE ; Exo:=VH^.Suivant ; END Else
         BEGIN
         For i:=1 To 5 Do BEGIN If CodeExo=VH^.ExoClo[i].Code then
            BEGIN Exo:=VH^.ExoClo[i] ; Result:=TRUE ; Exit ; END ; END ;
          // ajout me
              if VH^.Exercices[1].Code <> '' then
              begin
                    for  i:= 1  to length (VH^.Exercices) do
                    begin
                                if (VH^.Exercices[i].Code = CodeExo) then
                                begin
                                     Exo.Code    := VH^.Exercices[i].Code ;
                                     Exo.Deb     := VH^.Exercices[i].Deb ;
                                     Exo.Fin     := VH^.Exercices[i].Fin ;
                                     Result      := TRUE ;
                                     break;
                                end;
                    end;
              end
              else
              begin
                   Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="'+CodeExo+'"',TRUE,-1,'',true);
                   If Not Q.Eof Then
                      BEGIN
                      Exo.Code:=Q.FindField('EX_EXERCICE').AsString ;
                      Exo.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
                      Exo.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
                      Result:=TRUE ;
                      END ;
                   Ferme(Q) ;
              end;
        END ;
{$ENDIF NOVH}
END ;

{$IFNDEF NOVH}
Function LWhereV8 : String ;
BEGIN
Result:='' ;
if VH^.ExoV8.Code<>'' then Result:='E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ' ;
END ;

procedure ListeUnePeriode(Exo : String3 ; LLItems,LLValues : HTStrings ; Deb : Boolean; bAvecCloture : Boolean=TRUE) ;
{ Création des périodes comptable pour un exercice donné }
var DateExo      : TExoDate ;
    i            : integer ;
    AA,MM,NbMois : Word ;
    dd : TDateTime ;
    SDat : String ;
BEGIN
NbMois:=0 ;
QuelDateDeExo(Exo,DateExo) ;
If DateExo.Code<>'' Then
   BEGIN
   NOMBREPEREXO(DateExo,MM,AA,NbMois) ;
   for i:=0 to NbMois-1 do
       BEGIN
       dd:=PlusMois(DateExo.Deb,i) ;
       if (not bAvecCloture) and (VH^.DateCloturePer>0) and (FinDeMois(dd)<=VH^.DateCloturePer) then Continue ;
       SDat:=FormatDateTime('mmmm yyyy',dd) ;
       LLItems.add(FirstMajuscule(SDat)) ;
       If Deb Then LLValues.add(DateToStr(DebutdeMois(dd))) Else LLValues.add(DateToStr(FindeMois(dd)))
       END ;
   END ;
END ;

Procedure ListePeriode(Exo : String3 ; LLItems,LLValues : HTStrings ; Deb : Boolean; bAvecCloture : Boolean=TRUE) ;
{ Création des périodes comptable pour un ou tous les exercices }
BEGIN
LLItems.Clear ; LLValues.Clear ;
if Exo='' then
  BEGIN
  ListeUnePeriode(VH^.precedent.Code,LLItems,LLValues,Deb,bAvecCloture) ;
  ListeUnePeriode(VH^.EnCours.Code,LLItems,LLValues,Deb,bAvecCloture) ;
  ListeUnePeriode(VH^.Suivant.Code,LLItems,LLValues,Deb,bAvecCloture) ;
  END else ListeUnePeriode(Exo,LLItems,LLValues,Deb,bAvecCloture) ;
END ;
{$ENDIF}


function DateBudget(CPER : THValComboBox) : String ;
Var ii : integer ;
BEGIN
ii:=CPER.Items.IndexOf(CPER.Text) ;
if ii>=0 then BEGIN If ii<=CPER.Values.Count Then Result:=CPER.Values[ii] else Result:='' ; END ;
END ;

(*======================================================================*)
{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 21/07/2005
Modifié le ... : 29/11/2006
Description .. : - LG - 21/07/2005 - FB 12939 - ajout du raccourci pour le 
Suite ........ : lettrage
Suite ........ : - LG - 29/11/2006 - affectation  du shortcut sur le bouton 
Suite ........ : fermer
Mots clefs ... : 
*****************************************************************}
function AttribShortCut ( Nom : String ) : String ;
Var RS : TShortCut ;
BEGIN
Result:='' ; RS:=0 ;
Nom:=uppercase(Nom) ;
//if Nom='VENTIL'       then RS:=ShortCut(65,[ssAlt]) else
if Nom='VENTIL'       then RS:=ShortCut(65,[ssCtrl]) else
if Nom='MODIFECHE'    then RS:=ShortCut(66,[ssCtrl]) else
if Nom='LANCESAISIEBOR' then RS:=ShortCut(66,[ssAlt]) else
if Nom='COMPLEMENT'   then RS:=ShortCut(67,[ssAlt]) else
if Nom='INFOPOINTAGE' then RS:=ShortCut(67,[ssAlt]) else
if Nom='COMBI'        then RS:=ShortCut(67,[ssAlt]) else
if Nom='MONNAIE'      then RS:=ShortCut(68,[ssCtrl]) else
if Nom='ECHE'         then RS:=ShortCut(69,[ssAlt]) else
if Nom='CHERCHER'     then RS:=ShortCut(70,[ssCtrl]) else
if Nom='GUIDE'        then RS:=ShortCut(71,[ssAlt]) else
if Nom='CREERGUIDE'   then RS:=ShortCut(71,[ssCtrl]) else
if Nom='CHANCEL'      then RS:=ShortCut(72,[ssAlt]) else
if Nom='ZOOMDEVISE'   then RS:=ShortCut(73,[ssAlt]) else
if Nom='ZOOMJOURNAL'  then RS:=ShortCut(74,[ssAlt]) else
if Nom='PRORATA'      then RS:=ShortCut(75,[ssCtrl]) else
if Nom='REF'          then RS:=ShortCut(76,[ssAlt]) else
if Nom='LIBAUTO'      then RS:=ShortCut(76,[ssAlt]) else
if Nom='LISTEPIECES'  then RS:=ShortCut(76,[ssCtrl]) else
if Nom='COMPPIECE'    then RS:=ShortCut(76,[ssCtrl]) else
if Nom='MODEPAIE'     then RS:=ShortCut(77,[ssAlt]) else
if Nom='PIECEGC'      then RS:=ShortCut(77,[ssAlt]) else
if Nom='MODIFGENERE'  then RS:=ShortCut(77,[ssAlt]) else
if Nom='ZOOMIMMO'     then RS:=ShortCut(79,[ssCtrl]) else
if Nom='PIECEHISTO'   then RS:=ShortCut(79,[ssCtrl]) else
if Nom='IMPRIMER'     then RS:=ShortCut(80,[ssCtrl]) else
if Nom='ETATPOINTAGE' then RS:=ShortCut(80,[ssCtrl]) else
if Nom='DERNPIECES'   then RS:=ShortCut(80,[ssAlt]) else
if Nom='CLEREPART'    then RS:=ShortCut(82,[ssCtrl]) else
if Nom='ETATRAPPRO'   then RS:=ShortCut(82,[ssCtrl]) else
if Nom='CHOIXREGIME'  then RS:=ShortCut(82,[ssCtrl]) else
if Nom='REPART'       then RS:=ShortCut(82,[ssAlt]) else
if Nom='CLEFBU'       then RS:=ShortCut(82,[ssAlt]) else
if Nom='MODIFRIB'     then RS:=ShortCut(82,[ssAlt]) else
if Nom='SOUSTOT'      then RS:=ShortCut(83,[ssCtrl]) else
if Nom='SCENARIO'     then RS:=ShortCut(83,[ssAlt]) else
if Nom='SOLDEGS'      then RS:=ShortCut(83,[ssAlt]) else
if Nom='LANCESAISIE'  then RS:=ShortCut(83,[ssAlt]) else
if Nom='ZOOMETABL'    then RS:=ShortCut(84,[ssAlt]) else
if Nom='ZOOMETABL'    then RS:=ShortCut(84,[ssCtrl]) else
if Nom='VENTILTYPE'   then RS:=ShortCut(86,[ssAlt]) else
if Nom='MODIFTVA'     then RS:=ShortCut(86,[ssAlt]) else
if Nom='VIDEPIECE'    then RS:=ShortCut(89,[ssCtrl]) else
if Nom='ANNULERSEL'   then RS:=ShortCut(90,[ssCtrl]) else
if Nom='APPLIQUER'    then RS:=ShortCut(VK_RETURN,[ssCtrl]) else
if Nom='POINTE'       then RS:=ShortCut(VK_RETURN,[]) else
if Nom='ABANDON'      then RS:=ShortCut(VK_ESCAPE,[]) else
if Nom='CABANDON'     then RS:=ShortCut(VK_ESCAPE,[]) else
if Nom='INSERT'       then RS:=ShortCut(VK_INSERT,[]) else
if Nom='RAJOUTE'      then RS:=ShortCut(VK_INSERT,[ssShift]) else
if Nom='DEL'          then RS:=ShortCut(VK_DELETE,[]) else
if Nom='SDEL'         then RS:=ShortCut(VK_DELETE,[ssCtrl]) else
if Nom='ENLEVE'       then RS:=ShortCut(VK_DELETE,[ssShift]) else
if Nom='AIDE'         then RS:=ShortCut(VK_F1,[]) else
if Nom='CAIDE'        then RS:=ShortCut(VK_F1,[]) else
if Nom='PREV'         then RS:=ShortCut(VK_F3,[]) else
if Nom='FIRST'        then RS:=ShortCut(VK_F3,[ssShift]) else
if Nom='NEXT'         then RS:=ShortCut(VK_F4,[]) else
if Nom='LAST'         then RS:=ShortCut(VK_F4,[ssShift]) else
if Nom='ZOOM'         then RS:=ShortCut(VK_F5,[]) else
if Nom='ZOOMECRITURE' then RS:=ShortCut(VK_F5,[]) else
if Nom='ZOOMBUDGET'   then RS:=ShortCut(VK_F5,[]) else
if Nom='ASSISTANT'    then RS:=ShortCut(VK_F5,[]) else
if Nom='ZOOMABO'      then RS:=ShortCut(VK_F5,[ssShift]) else
if Nom='ZOOMPIECE'    then RS:=ShortCut(VK_F5,[ssShift]) else
if Nom='ZOOMCPTE'     then RS:=ShortCut(VK_F5,[ssShift]) else
if Nom='ZOOMSECTION'  then RS:=ShortCut(VK_F5,[ssAlt]) else
if Nom='SOLDE'        then RS:=ShortCut(VK_F6,[]) else
if Nom='ZOOMMVTEFF'   then RS:=ShortCut(VK_F6,[]) else
if Nom='SWAPPIVOT'    then RS:=ShortCut(VK_F8,[ssAlt]) else
if Nom='SWAPEURO'     then RS:=ShortCut(VK_F8,[ssShift]) else
if Nom='LIG'          then RS:=ShortCut(VK_F7,[]) else
if Nom='CONTROLETVA'  then RS:=ShortCut(VK_F8,[]) else
if Nom='GENERETVA'    then RS:=ShortCut(VK_F9,[]) else
if Nom='VALIDE'       then RS:=ShortCut(VK_F10,[]) else
if Nom='VALIDER'      then RS:=ShortCut(VK_F10,[]) else
if Nom='CVALIDE'      then RS:=ShortCut(VK_F10,[]) else
if Nom='VALIDESELECT' then RS:=ShortCut(VK_F10,[]) else
if Nom='FERME'        then RS:=ShortCut(VK_ESCAPE,[]) else
if Nom='LETTRAGE'     then RS:=ShortCut(Word('L'),[ssCtrl]) else
 ;
if RS<>0 then Result:=ShortCutToText(RS) ;
END ;
{$ENDIF}

Function AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
BEGIN
if OkSymbole then
   BEGIN
   if LeMontant=0 then AfficheMontant:=''
                  else AfficheMontant:=FormatFloat(Formatage+' '+LeSymbole,LeMontant) ;
   END else
   BEGIN
   if LeMontant=0 then AfficheMontant:=''
                  else AfficheMontant:=FormatFloat(Formatage,LeMontant) ;
   END ;
END ;

Procedure ChargeXuelib2 ;
BEGIN
END ;

Procedure SauveXuelib2 ;
BEGIN
END ;

{$IFNDEF NOVH}
Procedure VerifOldTeletrans ;
Var i : Integer ;
    St : String ;
BEGIN
If ParamCount>0 Then
  for i:=1 to ParamCount do
     BEGIN
     St:=ParamStr(i) ; If UpperCase(Trim(St))='TELETRANS' Then VH^.OldTeleTrans:=TRUE ;
     END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Procedure InitLaVariableHalley;
BEGIN
{$IFNDEF EAGLSERVER}
  New(VH) ;
	VH^.BStatutLiasse := nil;
  VH^.BStatutRevision := nil;
{$ENDIF !EAGLSERVER}
FillChar(VH^,Sizeof(VH^),#0) ;
VH^.ImportRL:=FALSE ; VH^.TenueEuro:=False ; VH^.Mugler:=FALSE ; VH^.MultiSouche:=FALSE ;
VH^.PasParSoc:=False ; VH^.BouclerSaisieCreat:=True ; VH^.AttribRIBAuto:=False ;
//VH^.SpeedCum:=FALSE ;
VH^.TOBCum:=Tob.create('Maman',Nil,-1) ;
VH^.UseTC:=(VH^.PaysLocalisation=CodeISOES) ; //XMG 31/07/03
(*======= PFUGIER *)
VH^.bByGenEtat:=FALSE ; //VH^.SpeedObj:=nil ; VH^.bSpeedFree:=FALSE ;
(*======= PFUGIER *)
{$IFDEF SPEC302}
VH^.PasParSoc:=TRUE ;
VH^.DefCatTVA:='TVA' ; VH^.DefCatTPF:='TPF' ;
{$ELSE}
VH^.DefCatTVA:='TX1' ; VH^.DefCatTPF:='TX2' ;
VH^.LaTOBTVA:=TOB.Create('',Nil,-1) ;
VH^.TOBJalEffet:=TOB.Create('',Nil,-1) ;
VH^.OBImmo:=TOB.Create('',Nil,-1) ;
VH^.ChargeOBImmo := True;
VH^.MPACC:=TStringList.Create ;
{$ENDIF}
ChargeXuelib2 ;
VerifOldTeleTrans ;
// GCO - 25/04/2002
VH^.EnSerie := False;
// FIN GCO
END ;
{$ENDIF !NOVH}

{=====================================================================}
{$IFNDEF NOVH}
{$IFNDEF EAGLSERVER}
Procedure FreeListeSP ;
Var fb : TFichierBase ;
    i : Integer ;
BEGIN
For fb:=fbAxe1 to fbAxe5 Do For i:=1 To MaxSousPlan Do If VH^.SousPlanAxe[fb,i].ListeSP<>NIL Then
  BEGIN
  VH^.SousPlanAxe[fb,i].ListeSP.Clear ; VH^.SousPlanAxe[fb,i].ListeSP.Free ;
  END ;
END ;
{$ELSE}
Procedure FreeListeSP(RelpVH:PLaVariableHalley) ;
Var fb : TFichierBase ;
    i : Integer ;
BEGIN
For fb:=fbAxe1 to fbAxe5 Do For i:=1 To MaxSousPlan Do If RelpVH^.SousPlanAxe[fb,i].ListeSP<>NIL Then
  BEGIN
  RelpVH^.SousPlanAxe[fb,i].ListeSP.Clear ; RelpVH^.SousPlanAxe[fb,i].ListeSP.Free ;
  END ;
END ;
{$ENDIF}
{$ENDIF !NOVH}

{=====================================================================}
{$IFNDEF NOVH}
{$IFNDEF HVCL}

{$IFNDEF EAGLSERVER}
Procedure ReleaseLaVariableHalley ();
begin
{Compta}
if VH <> nil then
  begin
  VH^.LaTOBTVA.Free ;
  VH^.TOBJalEffet.Free ;
  VH^.OBImmo.Free ;
  VH^.MPACC.Clear ; VH^.MPACC.Free ;
  VH^.TOBCUM.Free ;
  FreeListeSP ;
  Dispose(VH ) ;
  end;
end;
{$ELSE}
Procedure ReleaseLaVariableHalley (RelpVH:PLaVariableHalley);
//Var HH : TextFile ;
begin
{Compta}
(*
If GMM<50 Then
  BEGIN
  Assign(HH,'c:\GG'+IntToStr(GMM)+'.TXT') ;
  Rewrite(HH) ;
  Writeln(HH,'VAS Y GASTON') ;
  Close(HH) ;
  Inc(GMM) ;
  END ;
*)
if RelpVH <> nil then
  begin
  RelpVH^.LaTOBTVA.Free ;
  RelpVH^.TOBJalEffet.Free ;
  RelpVH^.OBImmo.Free ;
  RelpVH^.MPACC.Clear ; RelpVH^.MPACC.Free ;
  RelpVH^.TOBCUM.Free ;
  FreeListeSP(RelpVH) ;
  end;
end;
{$ENDIF}



(*
Procedure ReleaseLaVariableHalley;
{$IFDEF EAGLSERVER}
Var HH : TextFile ;
{$ENDIF EAGLSERVER}
begin
  {Compta}
{$IFDEF EAGLSERVER}
If GGM<10 Then
  BEGIN
  Assign(HH,'c:\GG'+IntToStr(GMM)+'.TXT') ;
  Rewrite(HH) ;
  Writeln(HH,'VAS Y GASTON') ;
  Close(HH) ;
  Inc(GMM) ;
  END ;
{$ENDIF EAGLSERVER}
  if VH <> nil then
  begin
    VH^.LaTOBTVA.Free ;
    VH^.TOBJalEffet.Free ;
    VH^.OBImmo.Free ;
    VH^.MPACC.Clear ; VH^.MPACC.Free ;
    VH^.TOBCUM.Free ;
    FreeListeSP ;
    {$IFNDEF EAGLSERVER}
      Dispose(VH) ;
    {$ENDIF !EAGLSERVER}
  end;
end;
*)
{$ENDIF !HVCL}
{$ENDIF !NOVH}

{=====================================================================}
Function BourreLess ( St : String ; LeType : TFichierBase ) : string ;
var Lg,ll,i : Integer ;
    Bourre  : Char ;
    lInfoCpta : TInfoCpta ;
    Label 0 ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','BourreLess') ;
{$ENDIF}
lInfoCpta := GetInfoCpta(leType) ;
Lg := lInfoCpta.lg ; Bourre:=lInfoCpta.Cb ;
ll:=Length(St) ; If ll>Lg Then St:=Copy(St,1,Lg) ; i:=Length(St) ;
While i>0 Do BEGIN If St[i]=Bourre Then St[i]:=' ' Else Goto 0 ; Dec(i) ; END ;
0:Result:=Trim(St) ;
end ;


{=====================================================================}
Function BourreLaDonc ( St : String ; LeType : TFichierBase ) : string ;
var Lg,ll,i : Integer ;
    Bourre  : Char ;
    lInfoCpta : TInfoCpta ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','BourreLaDonc') ;
{$ENDIF}
If LeType In [fbAxe1..fbAux,fbNatCpt] Then
   BEGIN
   lInfoCpta := GetInfoCpta(leType) ;
   Lg := lInfoCpta.lg ;
   If LeType=fbNatCpt Then Bourre:=GetBourreLibre Else Bourre:=lInfoCpta.Cb ;
   Result:=St ; ll:=Length(Result) ;
   If ll<Lg then for i:=ll+1 to Lg do Result:=Result+Bourre ;
   END Else Result:=St ;
end ;
{=====================================================================}
Function BourreEtLess ( St : String ; LeType : TFichierBase ; ForceLg : Integer = 0) : string ;
var Lg,ll,i : Integer ;
    Bourre  : Char ;
     lInfoCpta : TInfoCpta ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','BourreEtLess') ;
{$ENDIF}
If LeType In [fbAxe1..fbAux,fbNatCpt] Then
   BEGIN
   lInfoCpta := GetInfoCpta(leType) ;
   Lg := lInfoCpta.lg ; If ForceLg>0 Then Lg:=ForceLg ;
   If LeType=fbNatCpt Then Bourre:=GetBourreLibre Else Bourre:=lInfoCpta.Cb ;
   If Length(St)>Lg Then St:=Trim(Copy(St,1,Lg)) ;
   Result:=St ; ll:=Length(Result) ;
   If ll<Lg then
      BEGIN
      for i:=ll+1 to Lg do Result:=Result+Bourre ;
      END Else Result:=Copy(Result,1,Lg) ;
   END Else Result:=St ;
end ;

function BourreOuTronque(St : String ; fb : TFichierBase) : String ;
  function LongMax2 ( fb : TFichierBase ) : integer ;
  Var i : Integer ;
  BEGIN
  if fb in [fbAxe1..fbAux] then Result:=GetInfoCpta(fb).lg else
  Case fb of
    fbJal     : Result:=3 ;
    fbAxe1SP1..fbAxe5SP6 : BEGIN      { DONE -olaurent -cdll serveur : a reecrire }
                           i:=((Byte(fb)-Byte(fbAxe1SP1)) Mod 6)+1 ;
                           Result:=GetSousPlanAxe(fbSousPlanTofb(fb),i).Longueur ;
                           END ;
    else Result:=17 ;
    end ;
  END ;

var Diff : Integer ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','BourreOuTronque') ;
{$ENDIF}
St:=Trim(St) ; Result:=St ; if (St='') then Exit ;
Diff:=Length(St)-LongMax2(fb) ;
if Diff>0 then St:=BourreLess(St,fb) ;
Result:=BourreLaDonc(St,fb) ;
END ;

{=====================================================================}
Function BourreLaDoncSurLaTable ( Code, St : String; ForceLg : Integer = 0) : string ;
var Lg,ll,i : Integer ;
{$IFNDEF EAGLSERVER}
 j : integer ;
{$ENDIF}
    Bourre    : Char ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','BourreLaDoncSurLaTable ') ;
{$ENDIF}
Result:=St ;
Case Code[1] Of
     'G' : i:=1 ; 'T' : i:=2 ; 'S' : i:=3 ; 'B' : i:=4 ;
     'D' : i:=5 ; 'E' : i:=6 ; 'A' : i:=7 ; 'U' : i:=8 ;
     'I' : i:=9 ;
     Else i:=0 ;
   END ;
{$IFDEF EAGLSERVER}
If i>0 Then Lg:=ForceLg Else exit ;
{$ELSE}
{$IFNDEF NOVH}
{ DONE -olaurent -cdll serveur : a faire }
If i>0 Then BEGIN j:=StrToInt(Copy(Code,2,2))+1 ; Lg:=GetLgTableLibre(i,j) ; END Else exit ;
{$ENDIF}
{$ENDIF}
Bourre:=GetBourreLibre ;
ll:=Length(Result) ;
If ll<Lg then for i:=ll+1 to Lg do Result:=Result+Bourre ;
end ;

function BourreLaDoncSurLesComptes(st : string ; c : string = '') : string;
var st1 : string ;
    l : integer;
    cb : string;
    lInfoCpta : TInfoCpta ;
begin
  lInfoCpta := GetInfoCpta(fbGene) ;
  if c = '' then cb := lInfoCpta.Cb else cb := c;
  st1 := Trim(st) ;
  l := lInfoCpta.Lg;
  if l<= 0 then {!!!!????} else
  if l<Length(st1) then st1 := Copy(st1,1,l);
  while Length(st1)<l do st1 := st1 + cb ;
  Result := st1 ;
end;


function CompteDansLeIntervalle(Quoi, Deb, Fin : string) : boolean;
var
  cptQuoi, cptDeb, cptFin : string;
begin
  cptQuoi := BourreLaDoncSurLesComptes(Quoi); // on bourre avec le caratère de VH^
  cptDeb := BourreLaDoncSurLesComptes(Deb,'0'); // on bourre avec des 0
  cptFin := BourreLaDoncSurLesComptes(Fin,'9'); // on bourre avec des 9
  result := (cptQuoi>=cptDeb) and (cptQuoi<=cptFin);
end;

////////////////////////////////////////////////////////////////////////////////////////////////
// FONCTIONS DE CONVERSIONS
Function NatureToTz (Nat : String) : TZoomTable ;
BEGIN
Case Nat[1] Of
    'B' : Result:=TZoomtable(Ord(tzNatBud0)+Ord(Nat[3])-48) ;
    'G' : Result:=TZoomtable(Ord(tzNatGene0)+Ord(Nat[3])-48) ;
    'S' : Result:=TZoomtable(Ord(tzNatSect0)+Ord(Nat[3])-48) ;
    'T' : Result:=TZoomtable(Ord(tzNatTiers0)+Ord(Nat[3])-48) ;
    'D' : Result:=TZoomtable(Ord(tzNatBudS0)+Ord(Nat[3])-48) ;
    'E' : Result:=TZoomtable(Ord(tzNatEcrE0)+Ord(Nat[3])-48) ;
    'A' : Result:=TZoomtable(Ord(tzNatEcrA0)+Ord(Nat[3])-48) ;
    'U' : Result:=TZoomtable(Ord(tzNatEcrU0)+Ord(Nat[3])-48) ;
    'I' : Result:=TZoomtable(Ord(tzNatImmo0)+Ord(Nat[3])-48) ;
    Else Result:=tzRien ;
  End ;
END ;

{JP 13/06/05 : Equivaut à NatureToTz pour les DataTypes
{---------------------------------------------------------------------------------------}
function NatureToDataType(CodeTable : string) : string;
{---------------------------------------------------------------------------------------}
begin
  case CodeTable[1] of
    'B' : Result:= 'TZNATBUD'   + CodeTable[3];
    'G' : Result:= 'TZNATGENE'  + CodeTable[3];
    'S' : Result:= 'TZNATSECT'  + CodeTable[3];
    'T' : Result:= 'TZNATTIERS' + CodeTable[3];
    'D' : Result:= 'TZNATBUDS'  + CodeTable[3];
    'E' : Result:= 'TZNATECR'   + CodeTable[3];
    'A' : Result:= 'TZNATECRA'  + CodeTable[3];
    'U' : Result:= 'TZNATECRU'  + CodeTable[3];
    'I' : Result:= 'TINATIMMO'  + CodeTable[3];
  end;
end;

{JP 13/06/05 : Equivaut à AxeToTz pour les DataTypes
{---------------------------------------------------------------------------------------}
function AxeToDataType(CodeAxe : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'TZSECTION' + CodeAxe[2];
end;

Function  AxeToTz ( Ax : String ) : TZoomTable ;
BEGIN
Result:=TZoomTable(Ord(tzSection)+Ord(Ax[2])-49) ;
END ;

Function StringToTz(s : String) : TZoomTable ;
begin
Result:=tzGeneral ;
if (s='GENERAL') or (s='GENERAUX') then result := tzGeneral else
if (s='AUXILIAIRE') or (s='TIERS') then result := tzTiers  else
if (s='SECTION') or (s='SECTION1') or (s='CSECTION1') then result := tzSection else
if (s='JOURNAL') or (s='JOURNAUX') then result := tzJournal ;
//vt à compléter
end ;

Function  tzToNature( tz : TZoomTable ) : String ;
BEGIN
Case tz of
  tzNatGene0,tzNatGene1,tzNatGene2,tzNatGene3,tzNatGene4,tzNatGene5,
  tzNatGene6,tzNatGene7,tzNatGene8,tzNatGene9 : Result:='G'+FormatFloat('00',Byte(tz)-Byte(tzNatGene0)) ;
  tzNatTiers0,tzNatTiers1,tzNatTiers2,tzNatTiers3,tzNatTiers4,tzNatTiers5,
  tzNatTiers6,tzNatTiers7,tzNatTiers8,tzNatTiers9 : Result:='T'+FormatFloat('00',Byte(tz)-Byte(tzNatTiers0)) ;
  tzNatSect0,tzNatSect1,tzNatSect2,tzNatSect3,tzNatSect4,tzNatSect5,
  tzNatSect6,tzNatSect7,tzNatSect8,tzNatSect9 : Result:='S'+FormatFloat('00',Byte(tz)-Byte(tzNatSect0)) ;
  tzNatBud0,tzNatBud1,tzNatBud2,tzNatBud3,tzNatBud4,tzNatBud5,
  tzNatBud6,tzNatBud7,tzNatBud8,tzNatBud9 : Result:='B'+FormatFloat('00',Byte(tz)-Byte(tzNatBud0)) ;
  tzNatBudS0,tzNatBudS1,tzNatBudS2,tzNatBudS3,tzNatBudS4,tzNatBudS5,
  tzNatBudS6,tzNatBudS7,tzNatBudS8,tzNatBudS9 : Result:='D'+FormatFloat('00',Byte(tz)-Byte(tzNatBudS0)) ;
  tzNatEcrE0,tzNatEcrE1,tzNatEcrE2,tzNatEcrE3 : Result:='E'+FormatFloat('00',Byte(tz)-Byte(tzNatEcrE0)) ;
  tzNatEcrA0,tzNatEcrA1,tzNatEcrA2,tzNatEcrA3 : Result:='A'+FormatFloat('00',Byte(tz)-Byte(tzNatEcrA0)) ;
  tzNatEcrU0,tzNatEcrU1,tzNatEcrU2,tzNatEcrU3 : Result:='U'+FormatFloat('00',Byte(tz)-Byte(tzNatEcrU0)) ;
  tzNatImmo0,tzNatImmo1,tzNatImmo2,tzNatImmo3,tzNatImmo4,tzNatImmo5,
  tzNatImmo6,tzNatImmo7,tzNatImmo8,tzNatImmo9 : Result:='I'+FormatFloat('00',Byte(tz)-Byte(tzNatImmo0)) ;
  END ;
END ;

Function  tzToChampNature( tz : TZoomTable ; AvecPrefixe : Boolean) : String ;
Var StPref,StNum : String ;

BEGIN
StPref:='' ; STNum:='0' ;
Case tz of
  tzNatGene0,tzNatGene1,tzNatGene2,tzNatGene3,tzNatGene4,tzNatGene5,
  tzNatGene6,tzNatGene7,tzNatGene8,tzNatGene9 : BEGIN StPref:='G_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatGene0)) ; END ;
  tzNatTiers0,tzNatTiers1,tzNatTiers2,tzNatTiers3,tzNatTiers4,tzNatTiers5,
  tzNatTiers6,tzNatTiers7,tzNatTiers8,tzNatTiers9 : BEGIN StPref:='T_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatTiers0)) ; END ;
  tzNatSect0,tzNatSect1,tzNatSect2,tzNatSect3,tzNatSect4,tzNatSect5,
  tzNatSect6,tzNatSect7,tzNatSect8,tzNatSect9 : BEGIN StPref:='S_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatSect0)) ; END ;
  tzNatBud0,tzNatBud1,tzNatBud2,tzNatBud3,tzNatBud4,tzNatBud5,
  tzNatBud6,tzNatBud7,tzNatBud8,tzNatBud9 : BEGIN StPref:='B_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatBud0)) ; END ;
  tzNatBudS0,tzNatBudS1,tzNatBudS2,tzNatBudS3,tzNatBudS4,tzNatBudS5,
  tzNatBudS6,tzNatBudS7,tzNatBudS8,tzNatBudS9 : BEGIN StPref:='D_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatBudS0)) ; END ;
  tzNatEcrE0,tzNatEcrE1,tzNatEcrE2,tzNatEcrE3 : BEGIN StPref:='E_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatEcrE0)) ; END ;
  tzNatEcrA0,tzNatEcrA1,tzNatEcrA2,tzNatEcrA3 : BEGIN StPref:='Y_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatEcrA0)) ; END ;
  tzNatEcrU0,tzNatEcrU1,tzNatEcrU2,tzNatEcrU3 : BEGIN StPref:='U_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatEcrU0)) ; END ;
  tzNatImmo0,tzNatImmo1,tzNatImmo2,tzNatImmo3,tzNatImmo4,tzNatImmo5,
  tzNatImmo6,tzNatImmo7,tzNatImmo8,tzNatImmo9 : BEGIN StPref:='I_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatImmo0)) ; END ;  tzRien : BEGIN Result:='' ; Exit ; END ;
  END ;
If Not AvecPrefixe Then StPref:='' ;
Result:=StPref+'TABLE'+StNum ;
END ;

Function  AxeToFb ( Axe : String ) : TFichierBase ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','AxeToFb') ;
{$ENDIF}
Result:=fbAxe1 ;
if Length(Axe)>=2 then Result:=TFichierBase(Ord(Axe[2])-49) ;
END ;

Function  AxeToFbBud ( Axe : String ) : TFichierBase ;
BEGIN
Result:=fbBudsec1 ;
if Length(Axe)>=2 then
   Case Axe[2] of
        '1' : Result:=fbBudsec1 ;
        '2' : Result:=fbBudsec2;
        '3' : Result:=fbBudsec3 ;
        '4' : Result:=fbBudsec4 ;
        '5' : Result:=fbBudsec5 ;
       End ;
END ;

Function  FbToAxe ( fb : TFichierBase ) : String ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','FbToAxe') ;
{$ENDIF}
if fb in [fbBudSec1] then Result:='A1' else
   if fb in [fbBudSec2] then Result:='A2' else
      if fb in [fbBudSec3] then Result:='A3' else
         if fb in [fbBudSec4] then Result:='A4' else
            if fb in [fbBudSec5] then Result:='A5' else Result:='A'+Chr(49+Byte(fb)) ;
END ;

Function FbSousPlantoFb ( fb : TFichierBase ) : TFichierBase ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','FbSousPlantoFb') ;
{$ENDIF}
Result:=fbAxe1 ;
Case fb Of
  fbAxe1SP1..fbAxe1SP6 : Result:=fbAxe1 ;
  fbAxe2SP1..fbAxe2SP6 : Result:=fbAxe2 ;
  fbAxe3SP1..fbAxe3SP6 : Result:=fbAxe3 ;
  fbAxe4SP1..fbAxe4SP6 : Result:=fbAxe4 ;
  fbAxe5SP1..fbAxe5SP6 : Result:=fbAxe5 ;
  END ;
END ;

// Récupère une action à partir d'une chaine
function StrToTA(psz : String) : TActionFiche;
begin
    if (psz = 'TACREATE') or (psz='ACTION=CREATION')                    then result := taCreat
    else if (psz = 'TACREATEENSERIE') or (psz='ACTION=CREATIONENSERIE') then result := taCreatEnSerie
    else if (psz = 'TACREATONE')                                        then result := taCreatOne
    else if (psz = 'TAMODIF') or (psz='ACTION=MODIFICATION')            then result := taModif
    else if (psz = 'TACONSULT') or (psz='ACTION=CONSULTATION')          then result := taConsult
    else if (psz = 'TAMODIFENSERIE')                                    then result := taModifEnSerie
    else if (psz = 'TADUPLIQUE')                                        then result := taDuplique
    else                                                                     result := taConsult;
end;

// Récupère une chaine à partir d'un mode d'action
function TAToStr(TA : TActionFiche) : String;
begin
    if (TA = taCreat)             then result := 'TACREATE'
    else if (TA = taCreatEnSerie) then result := 'TACREATEENSERIE'
    else if (TA = taCreatOne)     then result := 'TACREATONE'
    else if (TA = taModif)        then result := 'TAMODIF'
    else if (TA = taConsult)      then result := 'TACONSULT'
    else if (TA = taModifEnSerie) then result := 'TAMODIFENSERIE'
    else if (TA = taDuplique)     then result := 'TADUPLIQUE'
    else                               result := 'TACONSULT';
end;

{============================================================================}
PROCEDURE NOMBREPEREXO (Exo : TExoDate ; Var PremMois,PremAnnee,NombreMois : Word) ;
Var PremJour,DernAnnee,DernMois,DernJour : Word ;
BEGIN
if Exo.Deb>Exo.Fin then BEGIN NombreMois:=0 ; exit ; END ;
DecodeDate(Exo.Deb,PremAnnee,PremMois,PremJour) ;
DecodeDate(Exo.Fin,DernAnnee,DernMois,DernJour) ;
NombreMois:=12*(DernAnnee-PremAnnee)+(DernMois-PremMois+1) ;
END ;

procedure AvertirMultiTable (FTypeTable : String ) ;
begin
  FTypeTable:=TRIM(UPPERCASE(FTypeTable)) ;
  if FTypeTable='TTJOURNAL' then
  begin
    AvertirTable('ttJournal') ;
    AvertirTable('ttJalSaisie') ;
    AvertirTable('ttJalSansEcart') ;
    AvertirTable('ttJalBanque') ;
    AvertirTable('ttJalNonANouveau') ;
    AvertirTable('ttJalOD') ;
    AvertirTable('ttJalAnalytique') ;
    AvertirTable('ttJalRegul') ;
    AvertirTable('ttJalRegulDevise') ;
    AvertirTable('ttJalEcart') ;
    AvertirTable('ttJalNonBanque') ;
    AvertirTable('ttJalANouveau') ;
    AvertirTable('ttJalTVA') ;
    AvertirTable('ttJalGuide') ;
    AvertirTable('ttJalAnalAN') ;
    AvertirTable('ttJalVENOD') ;
    AvertirTable('ttBudJalSansCAT') ;
    AvertirTable('ttJALEFFET') ;
    AvertirTable('ttJALFOLIO') ;
    AvertirTable('ttJOURNAUX') ;
    AvertirTable('ttJALSCENARIO') ;
    AvertirTable('CPJOURNALIFRS');
  end
  else
  if FTypeTable='TTCATJALBUD' then
  begin
    AvertirTable('ttCatJalBud') ;
    AvertirTable('ttCatJalBud1') ;
    AvertirTable('ttCatJalBud2') ;
    AvertirTable('ttCatJalBud3') ;
    AvertirTable('ttCatJalBud4') ;
    AvertirTable('ttCatJalBud5') ;
  end
  else
  // GCO - 20/09/2007 - FQ 21474 + Maj de toutes les tablettes sur EXERCICE
  if FTypeTable = 'TTEXERCICE' then
  begin
    AvertirTable('TTEXERCICE');
    AvertirTable('TTEXERCICEAPURGER');
    AvertirTable('TTEXERCICEBUDGET');
    AvertirTable('TTEXERCICEOUV');
    AvertirTable('TTEXERCICEOUVCPR');
    AvertirTable('TTEXERCICETRIE');
  end
  else AvertirTable(FTypeTable) ;
END ;

{=====================================================================}
function  RechargeParamSoc : boolean ;
BEGIN
{ DONE -olaurent -cdll serveur : a reecrire }
{$IFNDEF NOVH}
Result:=ChargeSocieteHalley ;
if Result then Result:=CHARGEMAGEXO(False) ;
{$ELSE}
result := false ;
{$ENDIF}
END ;

{$IFNDEF NOVH}
function CHARGEMAGEXO ( Parle : boolean ) : boolean ;
Var Premier : boolean ;
    Q       : TQuery ;
    DDeb,DFin : TDateTime ;
    SCode,SEtat : String3 ;
    i,j,k : Word ;
    ll,NbExoClo,Dep,mm  : Integer ;
    Exoavant : Array[1..20] Of Boolean ;
    OkOk : Boolean ;
    lStExoRef : string;
BEGIN
  Result:=True ;
  FillChar(VH^.Exercices,SizeOF(VH^.Exercices),#0) ;
  { Chargement de la liste des exercices }
  Q:=OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE, -1,'EXERCICE',true) ;
  if not Q.Eof then
  begin
    Q.First ; ll:=1 ;
    while ((not Q.Eof) and (ll<=20)) do
    begin
      VH^.Exercices[ll].Code := Q.FindField('EX_EXERCICE').AsString ;
      VH^.Exercices[ll].Deb := Q.FindField('EX_DATEDEBUT').AsDateTime ;
      VH^.Exercices[ll].Fin := Q.FindField('EX_DATEFIN').AsDateTime;
      VH^.Exercices[ll].DateButoir:=Q.FindField('EX_DATECUM').AsDateTime ;
      VH^.Exercices[ll].DateButoirRub:=Q.FindField('EX_DATECUMRUB').AsDateTime ;
      VH^.Exercices[ll].DateButoirBud:=Q.FindField('EX_DATECUMBUD').AsDateTime ;
      VH^.Exercices[ll].DateButoirBudgete:=Q.FindField('EX_DATECUMBUDGET').AsDateTime ;
      NombrePerExo(VH^.Exercices[ll],i,j,k) ; VH^.Exercices[ll].NombrePeriode:=k ;
      VH^.Exercices[ll].EtatCpta:=Q.FindField('EX_ETATCPTA').AsString;
      Inc(ll) ;
      Q.Next;
    end;
    VH^.Exercices[ll].Code := '';
    VH^.Exercices[ll].Deb := iDate1900;
    VH^.Exercices[ll].Fin := iDate1900;
  end else
  begin
    VH^.Exercices[1].Code := '';
    VH^.Exercices[1].Deb := iDate1900;
    VH^.Exercices[1].Fin := iDate1900;
  end;
  Ferme(Q) ;

// Charge Exo pour Compta
Q:=OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE,-1,'EXERCICE',true) ;
Premier:=True ; NbExoClo:=0 ;
{le suivant peut ne pas exister}
VH^.Entree.Deb:=0 ; VH^.Entree.Fin:=0 ; VH^.Entree.Code:='' ;
VH^.Suivant.Deb:=0 ; VH^.Suivant.Fin:=0 ; VH^.Suivant.Code:='' ;
VH^.EnCours.Deb:=0 ; VH^.EnCours.Fin:=0 ; VH^.EnCours.Code:='' ;
VH^.Precedent.Deb:=0 ; VH^.Precedent.Fin:=0 ; VH^.Precedent.Code:='' ;
fillchar(ExoAvant,SizeOf(ExoAvant),#0) ;
For i:=1 To 4 Do With VH^.ExoClo[i] Do BEGIN Deb:=0 ; Fin:=0 ; Code:='' ; END ;
While Not Q.EOF do
   BEGIN
   DDeb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
   DFin:=Q.FindField('EX_DATEFIN').AsDateTime ;
   SCode:=Q.FindField('EX_EXERCICE').AsString ;
   SEtat:=Q.FindField('EX_ETATCPTA').AsString ;
   If (VH^.ExoV8.Code=SCode) Then
      BEGIN
      VH^.ExoV8.Deb:=DDeb ; VH^.ExoV8.Fin:=DDeb ;
      END ;
   if SEtat='CDE' then
      BEGIN
      VH^.Precedent.Deb:=DDeb ; VH^.Precedent.Fin:=DFin ; VH^.Precedent.Code:=SCode ;
      NombrePerExo(VH^.Precedent,i,j,k) ; VH^.Precedent.NombrePeriode:=k ;
      Inc(NbExoClo) ; ExoAvant[NbExoClo]:=TRUE ;
      END ;
   if (SEtat='OUV') Or (SEtat='CPR') then
      BEGIN
      if Premier then
         BEGIN
         VH^.Encours.Deb:=DDeb ; VH^.Encours.Fin:=DFin ; VH^.Encours.Code:=SCode ;
         VH^.EnCours.DateButoir:=Q.FindField('EX_DATECUM').AsDateTime ;
         VH^.EnCours.DateButoirRub:=Q.FindField('EX_DATECUMRUB').AsDateTime ;
         VH^.EnCours.DateButoirBud:=Q.FindField('EX_DATECUMBUD').AsDateTime ;
         VH^.EnCours.DateButoirBudgete:=Q.FindField('EX_DATECUMBUDGET').AsDateTime ;
         NombrePerExo(VH^.EnCours,i,j,k) ; VH^.EnCours.NombrePeriode:=k ;
         VH^.EnCours.EtatCpta:=SEtat ;
         END else
         BEGIN
         VH^.Suivant.Deb:=DDeb ; VH^.Suivant.Fin:=DFin ; VH^.Suivant.Code:=SCode ;
         VH^.Suivant.DateButoir:=Q.FindField('EX_DATECUM').AsDateTime ;
         VH^.Suivant.DateButoirRub:=Q.FindField('EX_DATECUMRUB').AsDateTime ;
         VH^.Suivant.DateButoirBud:=Q.FindField('EX_DATECUMBUD').AsDateTime ;
         VH^.Suivant.DateButoirBudgete:=Q.FindField('EX_DATECUMBUDGET').AsDateTime ;
         NombrePerExo(VH^.Suivant,i,j,k) ; VH^.Suivant.NombrePeriode:=k ;
         VH^.Suivant.EtatCpta:=SEtat ;
         Break ;
         END ;
      Premier:=False ;
      END ;
   Q.Next ;
   END ;
If NbExoClo>=1 Then
   BEGIN
   Dep:=1 ; If NbExoClo>5 Then Dep:=NbExoClo-5+1 ;
   ll:=0 ; mm:=0 ;
   Q.First ;
   While Not Q.EOF do
      BEGIN
      Inc(ll) ;
      If (ll>=Dep) And (ll<Dep+5) Then
         BEGIN
         DDeb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
         DFin:=Q.FindField('EX_DATEFIN').AsDateTime ;
         SCode:=Q.FindField('EX_EXERCICE').AsString ;
         SEtat:=Q.FindField('EX_ETATCPTA').AsString ;
         if SEtat='CDE' then
            BEGIN
            Inc(mm) ;
            VH^.ExoClo[mm].Deb:=DDeb ; VH^.ExoClo[mm].Fin:=DFin ; VH^.ExoClo[mm].Code:=SCode ;
            VH^.ExoClo[mm].DateButoir:=Q.FindField('EX_DATECUM').AsDateTime ;
            VH^.ExoClo[mm].DateButoirRub:=Q.FindField('EX_DATECUMRUB').AsDateTime ;
            NombrePerExo(VH^.ExoClo[mm],i,j,k) ; VH^.ExoClo[mm].NombrePeriode:=k ;
            END ;
         END ;
      Q.Next ;
      END ;
   END ;
Ferme(Q) ;
If (V_PGI.DateEntree>=VH^.Encours.Deb) And (V_PGI.DateEntree<=VH^.Encours.Fin) Then VH^.Entree:=VH^.Encours else
 If (V_PGI.DateEntree>=VH^.Suivant.Deb) And (V_PGI.DateEntree<=VH^.Suivant.Fin) Then VH^.Entree:=VH^.Suivant ;
if VH^.Entree.Code='' then
   BEGIN
   VH^.Entree:=VH^.Encours ;
   if ((V_PGI.Halley) and (Parle)) then Result:=False ;
   END ;
OkOk:=TRUE ;
{$IFNDEF SPEC350}
{$IFDEF ALLVERSION}
If VH^.VNUMSOC<500 Then OkOK:=FALSE ;
{$ENDIF}
If OkOk Then
BEGIN
  // GCO - 31/08/2006 - FQ 18521
  VH^.CPExoRef.Code := '';
  VH^.CPExoRef.Deb  := 0;
  VH^.CPExoRef.Fin  := 0;

  lStExoRef := GetParamSocSecur('SO_CPEXOREF', '', True);
  if lStExoRef <> '' then
  begin
    Q := OpenSQL('SELECT EX_EXERCICE, EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE ' +
                 'EX_EXERCICE = "' + lStExoRef + '"', True,-1,'',true);
    if not Q.Eof then
    begin
      VH^.CPExoRef.Code := Q.FindField('EX_EXERCICE').AsString;
      VH^.CPExoRef.Deb  := Q.FindField('EX_DATEDEBUT').AsDateTime;
      VH^.CPExoRef.Fin  := Q.FindField('EX_DATEFIN').AsDateTime;
    end
    else
    begin
      // Le code Exercice du ParamSOC n'existe pas en table EXERCICE (PB FQ 18521)
      // Correction du PB en associant l'encours au ParamSOC
      VH^.CPExoRef.Code := VH^.EnCours.Code;
      VH^.CPExoRef.Deb  := VH^.EnCours.Deb;
      VH^.CPExoRef.Fin  := VH^.EnCours.Fin;
      SetParamSoc('SO_CPEXOREF', VH^.EnCours.Code);
    end;
    Ferme(Q);
  end;

  (*
  if VH^.CPExoRef.Code = VH^.Encours.Code then
  begin
    VH^.CPExoRef.Deb := VH^.Encours.Deb ;
    VH^.CPExoRef.Fin := VH^.Encours.Fin ;
  end
  else
    if VH^.CPExoRef.Code = VH^.Suivant.Code then
    begin
      VH^.CPExoRef.Deb := VH^.Suivant.Deb;
      VH^.CPExoRef.Fin := VH^.Suivant.Fin;
    end
    else  // CA - 27/11/2001
    begin
      if VH^.CPExoRef.Code <>'' then
      begin
        lStExoRef := GetParamSocSecur('SO_CPEXOREF', '', True);
        if VH^.CpExoRef.Code <> lStExoRef then
            VH^.CpExoRef.Code := lStExoRef;

        Q := OpenSQL ('SELECT EX_DATEDEBUT,EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+VH^.CPExoRef.Code+'"',True);
        if not Q.Eof then
        begin
          VH^.CPExoRef.Deb := Q.FindField('EX_DATEDEBUT').AsDateTime;
          VH^.CPExoRef.Fin := Q.FindField('EX_DATEFIN').AsDateTime;
        end
        else // GCO - 31/08/2006 - FQ 18521
        begin

        end;

        Ferme (Q);
      end;
    end;

  if VH^.CPExoRef.Deb=0 then VH^.CPExoRef.Code:='' ; *)
END
else VH^.CPExoRef.Code:='' ;

{$ENDIF}
{$IFDEF AMORTISSEMENT}
if (VH^.OkModImmo) or (V_PGI.VersionDemo) then ChargeImoExo ;
{$ELSE}
 {$IFDEF PGIIMMO}
if (VH^.OkModImmo) or (V_PGI.VersionDemo) then ChargeImoExo ;
 {$ENDIF}
{$ENDIF}

{$IFNDEF HVCL} // XP 19.05.2006
  ReleaseCtxExercice; // GCO - 12/05/2006 - FQ 17705
{$ENDIF}
END ;
{$ENDIF}

{$IFNDEF EAGLSERVER}
function CreerLigPop ( B : TBitBtn ; Owner : TComponent ; ShortC : Boolean ; C : Char ) : TMenuItem ;
Var Nom,SC : String ;
    T   : TMenuItem ;
BEGIN
Nom:=B.Name ; System.Delete(Nom,1,1) ;
T:=TMenuItem.Create(Owner) ; T.Name:=C+Nom ;
T.Caption:=B.Hint ; T.OnClick:=B.OnClick ; T.Tag:=B.Tag ;
if ShortC then SC:=AttribShortCut(Nom) else SC:='' ;
if SC<>'' then T.Caption:=T.Caption+#9+SC ;
Result:=T ;
END ;

Function IsButtonPop ( CC : TComponent ) : boolean ;
BEGIN
  Result := ((CC.ClassType=TBitBtn) or (CC.ClassType=THBitBtn) or (CC.ClassType=TSpeedButton) or (CC.ClassType=TToolbarButton97));
END ;

procedure InitPopup( F : TForm) ;
Var i,LeTag,k : integer ;
    T,N : TMenuItem ;
    B,BT  : TBitBtn ;
    PP    : TPopupMenu ;
BEGIN
PP:=TPopupMenu(F.FindComponent('POPS')) ; if PP=Nil then Exit ;
{Purge des précédents items}
PurgePopup(PP) ;
{Création menu pop premier niveau}
for i:=0 to F.ComponentCount-1 do
 if ((F.Components[i].Tag>0) and (TControl(F.Components[i]).Parent.Enabled) and (TControl(F.Components[i]).Parent.Visible))
    then if IsButtonPop(F.Components[i]) then
    BEGIN
    B:=TBitBtn(F.Components[i]) ;
    if ((B.Enabled) and ((B.Visible) or (B.Tag>=1000))) then BEGIN T:=CreerLigPop(B,PP,True,'P') ; PP.Items.Add(T) ; END ;
    END ;
{Création 2ème niveau si bouton "menu zoom"}
for i:=0 to F.ComponentCount-1 do
    if ((F.Components[i].Tag<0) and (IsButtonPop(F.Components[i]))) then
    BEGIN
    B:=TBitBtn(F.Components[i]) ; if Not B.Enabled then Continue ;
    LeTag:=-F.Components[i].Tag ;
    T:=CreerLigPop(B,PP,False,'P') ; T.OnClick:=Nil ; PP.Items.Add(T) ;
    for k:=0 to F.ComponentCount-1 do
       if ((F.Components[k].Tag=LeTag) and (IsButtonPop(F.Components[k]))) then
       BEGIN
       BT:=TBitBtn(F.Components[k]) ;
       if BT.Enabled then BEGIN N:=CreerLigPop(BT,T,True,'P') ; T.Add(N) ; END ;
       END ;
    END ;
{$IFNDEF GIGI}
AddMenuPop(PP,'','') ;
{$ENDIF}
END ;

{ BPY 26/02/2003 }

// - modification du test pour savoir si le Component est affecté ....
//   car en eAGL on as des Components sans parent donc plantage ....
// - ajout aussi de la fin de la fonction ttabSheet ... maintenant
//   elle afiche un menu ;)
// - de plus j'ai recopié le test supplementaire du ttabSheet qui
//   n'existait pas dans ttabSheet ...

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... : 26/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure PopZoom97 ( BM : TToolbarButton97 ; POPZ : TPopupMenu ) ;
var
    i : integer;
    T : TMenuItem;
    B : TBitBtn;
    //P : TPoint;
    F : TForm;
    //LeMsg : TMsg;
    OkVisible : Boolean;
    TTB : TTabSheet;
begin
    PurgePopup(POPZ);
    F:=TForm(BM.Owner);

    for i:=0 to F.ComponentCount-1 do
    begin
        if (IsButtonPop(F.Components[i])) then
            if ((F.Components[i].Tag=-BM.Tag) and (TControl(F.Components[i]).Parent.Enabled) and (TControl(F.Components[i]).Parent.Visible)) then
            begin
                if (TControl(F.Components[i]).Parent is ttabSheet) then
                begin
                    TTB := TTabSheet(TControl(F.Components[i]).Parent);
                    OkVisible := TTB.TabVisible;
                end
                else OkVisible := TControl(F.Components[i]).Parent.Visible ;

                if OkVisible Then
                begin
                    B := TBitBtn(F.Components[i]) ;
                    if (B.Enabled) then
                    begin
                        T := CreerLigPop(B,POPZ,False,'Z');
                        POPZ.Items.Add(T);
                    end;
                end;
            end;
    end;

//    P.X := 0;
//    P.Y := BM.Height;
//    P := BM.ClientToScreen(P);
//    POPZ.PopUp(P.X,P.Y);
//    While PeekMessage(LeMsg,HWND(0),WM_MOUSEFIRST,WM_MOUSELAST,PM_REMOVE) do ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 26/02/2003
Modifié le ... : 26/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure PopZoom ( BM : TBitBtn ; POPZ : TPopupMenu ) ;
Var
    i : integer;
    T : TMenuItem;
    B : TBitBtn;
    P : TPoint;
    F : TForm;
    LeMsg : TMsg;
    OkVisible : Boolean;
    TTB : TTabSheet;
begin
    PurgePopup(POPZ);
    F := TForm(BM.Owner);

    for i:=0 to F.ComponentCount-1 do
    begin
        (*
        if ((F.Components[i].Tag=-BM.Tag) and (TControl(F.Components[i]).Parent.Enabled) and
        (TControl(F.Components[i]).Parent.Visible) and (IsButtonPop(F.Components[i]))) then
        *)
        if (IsButtonPop(F.Components[i])) then
            if ((F.Components[i].Tag = -BM.Tag) and (TControl(F.Components[i]).Parent.Enabled)) then
            begin
                if (TControl(F.Components[i]).Parent is ttabSheet) then
                begin
                    TTB := TTabSheet(TControl(F.Components[i]).Parent);
                    OkVisible := TTB.TabVisible;
                end
                else OkVisible := TControl(F.Components[i]).Parent.Visible ;

                if OkVisible Then
                begin
                    B := TBitBtn(F.Components[i]) ;
                    if (B.Enabled) then
                    begin
                        T := CreerLigPop(B,POPZ,False,'Z');
                        POPZ.Items.Add(T);
                    end;
                end;
            end;
    end;

    P.X := 0;
    P.Y := BM.Height;
    P := BM.ClientToScreen(P);
    POPZ.PopUp(P.X,P.Y);
    While PeekMessage(LeMsg,HWND(0),WM_MOUSEFIRST,WM_MOUSELAST,PM_REMOVE) do ;
END ;
{ Fin BPY}

procedure PurgePopup( PP : TPopupMenu ) ;
Var M,N : TMenuItem ;
BEGIN
if PP=Nil then Exit ;
if PP.Items.Count<=0 then Exit ;
While PP.Items.Count>0 do
   BEGIN
   M:=PP.Items[0] ;
   While M.Count>0 do BEGIN N:=M.Items[0] ; M.Remove(N) ; N.Free ; END ;
   PP.Items.Remove(M) ; M.Free ;
   END ;
END ;
{$ENDIF}

{$IFNDEF NOVH}
Function GetVH(s : string) : string ;
begin
if s='VH^.ENCOURS.CODE' then result:=VH^.Encours.Code else
if s='VH^.SUIVANT.CODE' then result:=VH^.Suivant.Code else
if s='VH^.PRECEDENT.CODE' then result:=VH^.Precedent.Code else
if s='VH^.ENCOURS.DEB' then result:=DateToStr(VH^.EnCours.Deb) else
if s='VH^.ENCOURS.FIN' then result:=DateToStr(VH^.EnCours.Fin) else
if s='VH^.SUIVANT.DEB' then result:=DateToStr(VH^.Suivant.Deb) else
if s='VH^.SUIVANT.FIN' then result:=DateToStr(VH^.Suivant.Fin) else
if s='VH^.PRECEDENT.DEB' then result:=DateToStr(VH^.Precedent.Deb) else
if s='VH^.PRECEDENT.FIN' then result:=DateToStr(VH^.Precedent.Fin) else
if s='VH^.JALATP' then result:=VH^.JalATP else
if s='VH^.EXOV8' then result:=VH^.ExoV8.Code else
if s='VH^.JALVTP' then result:=VH^.JalVTP ;
end ;
{$ENDIF}

{$IFNDEF EAGLSERVER}
Procedure LookLesDocks ( FF : TForm ) ;
Var i : integer ;
    TD : TDock97 ;
BEGIN
{$IFDEF AGL530F}
if V_PGI.MODETSE then Exit ;
{$ELSE}
//if V_PGI.Color16 then Exit ;
{$ENDIF}
for i:=0 to FF.ComponentCount-1 do if FF.Components[i] is TDock97 then
    BEGIN
    TD:=TDock97(FF.Components[i]) ;
    TD.BackGroundTransparent:=True ;
    LoadFond(TD.BackGround,2,TD) ;
    END ;
END ;
{$ENDIF EAGLSERVER}

Function EuroOK : Boolean ;
BEGIN
//Result:=(StrToInt(Copy(V_PGI.NumVersion,1,1))>=3) Or V_PGI.SAV ;
Result:=TRUE ;
END ;

Function FinFrancs(DD : tDateTime) : Boolean ;
BEGIN
//Result:=(StrToInt(Copy(V_PGI.NumVersion,1,1))>=3) Or V_PGI.SAV ;
Result:=TRUE ;
END ;

Procedure LibellesTableLibre ( PZ : TTabSheet ; NamL,NamH,Pref : String ) ;
Var LesLib : HTStringList ;
    i : Integer ;
    St : String ;
    Trouver : Boolean ;
    TL      : TLabel ;
    TH      : TEdit ;
    FF      : TForm ;
BEGIN
LesLib:=HTStringList.Create ; GetLibelleTableLibre(Pref,LesLib) ; Trouver:=False ;
FF:=TForm(PZ.Owner) ;
for i:=0 to LesLib.Count-1 do
    BEGIN
    St:=LesLib.Strings[i] ;
    TL:=TLabel(FF.FindComponent(NamL+IntToStr(i))) ;
    TH:=TEdit(FF.FindComponent(NamH+IntToStr(i))) ;
    if TL<>Nil then
       BEGIN
       TL.Caption:=ReadTokenSt(St) ;
       TL.Enabled:=(St='X') ;
       if TH<>Nil then TH.Enabled:=TL.Enabled ;
       if ((St='X') and (TL.Visible)) then Trouver:=True ;
       END ;
    END ;
LesLib.Free ;
PZ.TabVisible:=Trouver ;
END ;

Procedure analyseTableLibre(Var S1,S2 : String) ;
Var St,S,St1,St2 : String ;
    i : Integer ;
BEGIN
If (Pos('&',S1)<=0) And (Pos('|',S1)<=0) Then Exit ;
St:=S1 ; St1:='' ; St2:='' ; i:=0 ;
While St<>'' do
  BEGIN
  S:=ReadTokenStV(St) ; Inc(i) ;
  If i<=10 Then St1:=St1+S+',' Else St2:=St2+S+',' ;
  if S='' then Continue ;
  END ;
S1:=St1 ; S2:=St2 ;
END ;

Procedure TraiteWhereSup(SS,Champ,Code : String ; LgCode : Integer ; Var WhereSup : String) ;
Var i : Integer ;
    Operateur : String ;
BEGIN
i:=Pos(Code,SS) ;
If i>0 Then
   BEGIN
   If Code='&' Then Operateur:='AND' Else Operateur:='OR' ;
   if WhereSup<>'' then WhereSup:=WhereSup+' '+Operateur ;
   SS:=Copy(SS,2,Length(SS)-1) ;
   if Length(SS)<LgCode then WhereSup:=WhereSup+' ('+Champ+' Like "'+SS+'%") '
                        else WhereSup:=WhereSup+' ('+Champ+' ="'+SS+'") ' ;
   END ;
END ;

FUNCTION QuelTable(S,Pref,Operateur,Champ : String ; LgCode : Integer ; Var WhereTableLibre : String) : Boolean ;
Var OkOk,ATraiter,PasLesVides : Boolean ;
    i : Integer ;
    SS,Where,S1,WhereSupEt,WhereSupOu : String ;
BEGIN
Result:=FALSE ; If Pref='' Then Exit ;
OkOk:=Pos(',',s)>0 ;
If Not Okok Then Exit ;
If S[Length(S)]<>',' Then S:=S+',' ;
WhereTableLibre:='' ; Where:='' ; WhereSupEt:='' ; WhereSupOu:='' ;
S1:='' ; AnalyseTableLibre(S,S1) ;
For i:=1 To 10 Do
  BEGIN
  SS:=ReadTokenstV(S) ;
  ATRaiter:=(Pos('-',SS)<=0) And (Pos('#',SS)<=0) ;
  PasLesVides:=Pos('*',SS)>0 ;
  If ATraiter Then
     BEGIN
//     Where:='' ;
     If PasLesVides Then Where:=Where+' AND ('+Pref+'TABLE'+FormatFloat('0',i-1)+'<>"") '
                    Else Where:=Where+' AND ('+Pref+'TABLE'+FormatFloat('0',i-1)+Operateur+'"'+SS+'") ' ;
     END ;
  END ;
If S1<>'' Then
   BEGIN
   While S1<>'' Do
     BEGIN
     SS:=ReadTokenstV(S1) ;
     TraiteWhereSup(SS,Champ,'&',LgCode,WhereSupEt) ;
     TraiteWhereSup(SS,Champ,'|',LgCode,WhereSupOu) ;
     END ;
   END ;
If Where<>'' Then
   BEGIN
   Delete(Where,2,3) ;
   WhereTableLibre:='('+Where+')' ;
   If WhereSupEt<>'' Then WhereTableLibre:=WhereTableLibre+' AND ('+WhereSupEt+') ' ;
   If WhereSupOu<>'' Then WhereTableLibre:=WhereTableLibre+' OR ('+WhereSupOu+') ' ;
   Result:=TRUE ;
   END ;
END ;

Function AnalyseCompte(St : String ; fb : TFichierBase ; Exclu,OnTableLibre : Boolean ; RenvoieVide : Boolean = TRUE; DoDelete : Boolean = True; LesCptes : TStringList = nil) : String ; // 12102
Var i,LgCode : Byte ;
    S,Stemp,S1,S2 : String ;
    Champ,Swhere,Pref,SwhereTableLibre1,SwhereTableLibre2 : String ;
    OkTableLibre1,OkTableLibre2 : Boolean ;
    PosI : Integer ;
{$IFDEF EAGLSERVER}
  Q : TQuery;
{$ENDIF EAGLSERVER}
BEGIN
Case fb of
    fbGene : BEGIN Pref:='G_' ; Champ:='G_GENERAL' ; END ;
    fbAux  : BEGIN Pref:='T_' ; Champ:='T_AUXILIAIRE' ; END ;
    fbImmo : BEGIN Pref:='I_' ; Champ:='I_IMMO' ; END ;
    fbBudgen : BEGIN Pref:='BG_' ; Champ:='BG_BUDGENE'  ; END ;
    fbBudSec1..fbBudSec5 : BEGIN Pref:='BS_' ; Champ:='BS_BUDSECT' ; END ;
    fbBudJal : BEGIN Pref:='' ; Champ:='BJ_BUDJAL' ; END ;
    fbJal : BEGIN Pref:='' ; Champ:='J_JOURNAL' ; END ;
    fbAxe1..fbAxe5 : BEGIN Pref:='S_' ; Champ:='S_SECTION' ; END ;
    End ;
LgCode:=17 ;

{$IFDEF EAGLSERVER}
  if fb <= fbAux then begin
    {Pas de VH en EAGLSERVER}
    if fb = fbAux then LgCode := GetParamSocSecur('SO_LGCPTEAUX',0)
    else if fb = fbGene then LgCode := GetParamSocSecur('SO_LGCPTEGEN',0)
    else if fb in [fbAxe1, fbAxe2, fbAxe3, fbAxe4, fbAxe5] then begin
      case fb of
        fbAxe1 : Q := OpenSQL('SELECT X_LONGSECTION FROM AXE WHERE X_AXE = "A1" ORDER BY X_AXE', True, -1, 'LGSECTION1',true);
        fbAxe2 : Q := OpenSQL('SELECT X_LONGSECTION FROM AXE WHERE X_AXE = "A2" ORDER BY X_AXE', True, -1, 'LGSECTION2',true);
        fbAxe3 : Q := OpenSQL('SELECT X_LONGSECTION FROM AXE WHERE X_AXE = "A3" ORDER BY X_AXE', True, -1, 'LGSECTION3',true);
        fbAxe4 : Q := OpenSQL('SELECT X_LONGSECTION FROM AXE WHERE X_AXE = "A4" ORDER BY X_AXE', True, -1, 'LGSECTION4',true);
        fbAxe5 : Q := OpenSQL('SELECT X_LONGSECTION FROM AXE WHERE X_AXE = "A5" ORDER BY X_AXE', True, -1, 'LGSECTION5',true);
      end;
      if not Q.EOF then
        LgCode := Q.FindField('X_LONGSECTION').AsInteger;
      Ferme(Q) ;
    end;
  end;
{$ELSE}
if fb<=fbAux then LgCode:=VH^.Cpta[fb].Lg ;
{$ENDIF EAGLSERVER}

if((Champ='S_SECTION') Or (Champ='BS_BUDSECT')) And (Not Exclu) And (Pos('(',St)<=0) And DoDelete And
  ((Pos('A1',St)=1) or (Pos('A2',St)=1) or (Pos('A3',St)=1) or (Pos('A4',St)=1) or (Pos('A5',St)=1))then Delete(St,1,2) ;
Swhere:='' ;
While St<>'' do
  BEGIN
  S:=ReadTokenSt(St) ;
  if S='' then Continue ;
  i:=Pos('(',S) ;
  if i>0 then Stemp:=Copy(S,1,i-1) ;
  if (i<=0) And (Not Exclu) then Stemp:=S ;
  if Not Exclu then i:=Pos(':',Stemp)
               else i:=Pos(':',S) ;
  PosI:=Pos('?',STemp) ;
  if i>0 then
     BEGIN
     if Not Exclu then
        BEGIN
        S1:=Copy(Stemp,1,i-1) ; S2:=Copy(Stemp,i+1,Length(Stemp)) ;
        if Swhere<>'' then Swhere:=Swhere+' Or ' ;
        SWhereTableLibre1:='' ;
        OkTableLibre1:=OnTableLibre And QuelTable(S1,Pref,'>=',Champ,LgCode,SWhereTableLibre1) ;
        SWhereTableLibre2:='' ;
        OkTableLibre2:=OnTableLibre And QuelTable(S2,Pref,'<=',Champ,LgCode,SWhereTableLibre2) ;
        If OkTableLibre1 Or OkTableLibre2 Then
           BEGIN
           if SWhereTableLibre1<>'' then SWhere:=SWhere+SWhereTableLibre1 ;
           if SWhere<>'' then
              BEGIN
              if SWhereTableLibre2<>'' then SWhere:=SWhere+' AND '+SWhereTableLibre2 ;
              END else
              BEGIN
              if SWhereTableLibre2<>'' then SWhere:=SWhere+SWhereTableLibre2 ;
              END ;
           END Else
           BEGIN
           S2:=BourreLaDonc(S2,fb) ; S1:=BourreLaDonc(S1,fb) ;
           Swhere:=Swhere+' '+Champ+' >="'+S1+'" And '+Champ+' <="'+S2+'"' ;
           if LesCptes<> nil then LesCptes.Add (S1+':'+S2);
           END ;
        END else
        BEGIN
        S1:=Copy(S,1,i-1) ; S2:=Copy(S,i+1,Length(S)) ;
        S2:=BourreLaDonc(S2,fb) ; S1:=BourreLaDonc(S1,fb) ;
        if Swhere<>'' then Swhere:=Swhere+' And ' ;
        Swhere:=Swhere+' '+Champ+' Not BETWEEN "'+S1+'" And "'+S2+'"' ;
        if LesCptes<> nil then LesCptes.Add (S1+':'+S2);
        END ;
     END else
     BEGIN
     if Not Exclu then
        BEGIN
        if Swhere<>'' then Swhere:=Swhere+' Or ' ;
        SWhereTableLibre1:='' ;
        OkTableLibre1:=OnTableLibre And QuelTable(STemp,Pref,'=',Champ,LgCode,SWhereTableLibre1) ;
        If OkTableLibre1 Then SWhere:=SWhere+SWhereTableLibre1 Else
           BEGIN
           If PosI>0 Then
              BEGIN
              Stemp:=TraduitJoker(Stemp) ; Swhere:=Swhere+' '+Champ+' Like "'+Stemp+'"' ;
              if LesCptes<> nil then LesCptes.Add (Stemp);
              END Else
              BEGIN
              if Length(Stemp)<LgCode then
              begin
                Swhere:=Swhere+' '+Champ+' Like "'+Stemp+'%"';
                if LesCptes<> nil then LesCptes.Add (Stemp+'%');
              end else
              begin
                Swhere:=Swhere+' '+Champ+' ="'+Stemp+'"' ;
                if LesCptes<> nil then LesCptes.Add (Stemp);
              end;
              END ;
           END ;
        END else
        BEGIN
        if Swhere<>'' then Swhere:=Swhere+' And ' ;
        if Length(S)<LgCode then Swhere:=Swhere+' '+Champ+' Not Like "'+S+'%"'
                            else Swhere:=Swhere+' '+Champ+' <>"'+S+'"' ;
        if LesCptes<> nil then LesCptes.Add (S);
        END ;
     END ;
  END ;
If Not RenvoieVide Then If (SWhere='') And (Not Exclu) And (Not OnTableLibre) Then SWhere:=' '+Champ+'<>"'+W_W+'" ' ;
if(SWhere<>'')then Result:='( '+Swhere+' )'
              else Result:=SWhere ;
END ;

procedure UpdateCombosLookUp ;
{$IFNDEF EAGLSERVER}
Var i : integer ;
{$ENDIF !EAGLSERVER}
BEGIN
{$IFNDEF EAGLSERVER}
{$IFNDEF NOVH}
For i:=1 to High(V_PGI.DECombos) do
   BEGIN
   if ((*(Copy(V_PGI.DECombos[i].TT,1,3)='TTJ') and*) (V_PGI.DECombos[i].Prefixe='J') and
       (V_PGI.DECombos[i].MemLoad=ltMem)) then V_PGI.DECombos[i].SaisieCode:=VH^.JalLookUp ;
   if ((*(Copy(V_PGI.DECombos[i].TT,1,3)='TTE') and*) (V_PGI.DECombos[i].Prefixe='ET') and
       (V_PGI.DECombos[i].MemLoad=LtMem)) then V_PGI.DECombos[i].SaisieCode:=VH^.EtabLookUp ;
   END ;
{$ENDIF}
{$ENDIF EAGLSERVER}
END ;

{$IFNDEF EAGLSERVER}
{$IFNDEF NOVH}
procedure UpdateSeries ;
Var ii : integer ;
    St : String ;
BEGIN
  {JP 22/08/05 : FQ 16481 : Gestion de la langue sur les modèles d'état de la gestion des tiers / CCMP}
  St := 'TTMODELESAI'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELERLC'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELERELANCE'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELELETTREVIR'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELELETTRETRA'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELELETTREPRE'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELELETTRECHQ'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELEGFF'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELEBOR'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  St := 'TTMODELEBAP'; ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := V_PGI.DECombos[ii].Where + ' AND MO_LANGUE = "' + V_PGI.LanguePrinc + '"';
  {FIN JP 22/08/05 : FQ 16481}


if EstSerie(S5) then
    BEGIN
      {Axes}
      {St:='TTAXE' ; ii:=TTToNum(St) ; JP 02/08/05 : Cela n'a pas de raison d'être}
      {Qualifiants pièce}
      St:='TTQUALPIECECRIT' ; ii:=TTToNum(St) ;
      if ii>0 then
        V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"PRE" AND CO_CODE<>"TOU"';
      St:='TTQUALPIECE' ; ii:=TTToNum(St) ;
      if ii>0 then
      BEGIN
      if EstComptaIFRS then
         V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"P" AND CO_CODE<>"R"'
      else
          V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"P" AND CO_CODE<>"R" AND CO_CODE<>"I"';
      END ;
    END
 // Modif IFRS S5 pack avancé
 else
    begin
    St:='TTQUALPIECE' ; ii:=TTToNum(St) ;
    if ii>0 then
      if not EstComptaIFRS
        then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"I"' ;
    end ;
 // Fin Modif IFRS S5 pack avancé

END ;
{$ENDIF}
{$ENDIF}

Function GetPeriode ( LaDate : TDateTime ) : integer ;
Var YY,MM,DD : Word ;
BEGIN
//Result:=0 ;
DecodeDate(LaDate,YY,MM,DD) ;
Result:=100*YY+MM ;
END ;

Function IncPeriode ( Periode : integer ) : integer;
begin
  if System.Copy(IntToStr(Periode),5,2) = '12' then
     Result := Periode + 89
  else
     Result := Periode + 1;
end;

Function  GetDateTimeFromPeriode ( Periode : integer ) : TDateTime;
begin
  Result := EncodeDate(StrToInt(Copy(IntToStr(Periode),1,4)),StrToInt(Copy(IntToStr(Periode),5,2)),1);
end;

procedure DirDefault(Sauve : TSaveDialog ; FileName : String) ;
var j,i : integer ;
BEGIN
j:=Length(FileName);
for i:=Length(FileName) downto 1 do if FileName[i]='\' then BEGIN j:=i ; Break ; END ;
Sauve.InitialDir:=Copy(FileName,1,j) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 15/04/2004
Modifié le ... :   /  /
Description .. : Regarde la présence d'une fiche immobilisation dans la
Suite ........ : base
Mots clefs ... :
*****************************************************************}
function AuMoinsUneImmo : Boolean ;
begin
  Result := ExisteSQL('SELECT I_IMMO FROM IMMO');
end;

Function PieceSurFolio(Jal : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Result:=FALSE ;
Q:=OpenSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="'+Jal+'" ',TRUE,-1,'',true);
If Not Q.Eof Then Result:=(Q.Fields[0].AsString='BOR') Or (Q.Fields[0].AsString='LIB') ;
Ferme(Q) ;
END ;


procedure MajBeforeDispatch(Num: Integer(*=0*)) ;
begin
// LG 08/02/207 utiliser seulement ds QR.pas
end ;

Procedure AGLChargeMagHalley ( Parms : array of variant ; nb : integer ) ;
BEGIN
{$IFNDEF NOVH}
ChargeMagHalley ;
{$ENDIF}
END ;

Function EstImmoPGI : Boolean ;
BEGIN
(*
{$IFDEF GGIMMO}
If ctxImmo In V_PGI.PGIContexte Then Result:=TRUE Else Result:=FALSE ;
{$ELSE}
*)
 {$IFDEF PGIIMMO}
 Result:=TRUE ;
 {$ELSE}
 Result:=FALSE ;
 {$ENDIF}
(*
{$ENDIF}
*)
END ;

{Si la tréso est utilisée}
function EstComptaTreso : Boolean;
begin
  {$IFDEF TRESO}
  Result := True; {FQ 10184}
  {$ELSE IF}
    Result := {JP 20/12/04 : FQ 15049}
              not (ctxPCL in V_PGI.PGIContexte) and
              (( (ctxTreso in V_PGI.PGIContexte) or GetParamSocSecur('SO_MODETRESO',False) )
              {JP / CA 26/10/07 : la fonction dépend du contexte et non de la sérialisation
               or ( V_PGI.VersionDemo )});
  {$ENDIF TRESO}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 20/08/2003
Modifié le ... : 07/11/2003
Description .. : permet de savoir si on a une version de CCS5 en Pack
Suite ........ : avancé (anciennement S7)
Suite ........ :
Suite ........ : BPY le 07/11/2003 : ajout pour S7 ... return true
Mots clefs ... :
*****************************************************************}
Function EstComptaPackAvance : Boolean ;
BEGIN
  //Actif pour tous
  Result := true ;
END ;
{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 20/08/2003
Modifié le ... : 20/08/2003
Description .. : permet de savoir si on a une version de CCS3 Premium
Suite ........ : avec Option ou supérieur.
Suite ........ : Si option cad VH^.OkModAna=True
Suite ........ : Alors Result:=False
Suite ........ : Sinon Result:=True
Mots clefs ... : PREMIUM
*****************************************************************}
{$IFNDEF NOVH}
Function EstComptaSansAna : Boolean ;
BEGIN
//init, par défaut on a droit à l'analytique
Result := False ; //plus de S3 désormais...Donc on retourne toujours False...
END ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 04/05/2004
Modifié le ... :   /  /
Description .. : Retourne Vrai si le module IFRS est sérialisé.
Suite ........ :
Suite ........ : A FINALISER Avec le test sur la séria
Mots clefs ... :
*****************************************************************}
function EstComptaIFRS : Boolean;
begin
  result := ( ( (ctxCompta in V_PGI.PGIContexte) and ( VH^.OkModCPPackIFRS ) )
              or ( V_PGI.VersionDemo )
             ) and (VH^.PaysLocalisation<>CodeISOES) ; //XVI 24/02/2005 //XVI Ce modiule n'existe pas en Espagne.
end ;
{$ENDIF}



{$IFNDEF NOVH}
Procedure CreerDeviseTenue(LaMonnaie : String) ;
Var Q : TQuery ;
BEGIN
If LaMonnaie='F' Then
  BEGIN
  Q:=OPENSQL('SELECT * FROM DEVISE WHERE D_DEVISE="FRF" ',FALSE) ;
  If Q.Eof Then
    BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('D_DEVISE').AsString:='FRF' ;
    Q.FindField('D_LIBELLE').AsString:='Francs' ;
    Q.FindField('D_SYMBOLE').AsString:='F' ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.FindField('D_DECIMALE').AsInteger:=2 ;
    Q.FindField('D_QUOTITE').AsInteger:=1 ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END Else
    BEGIN
    Q.Edit ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END ;
  Ferme(Q) ;
  ExecuteSQL('DELETE FROM DEVISE WHERE D_DEVISE="EUR"') ;
  SetParamSoc('SO_TENUEEURO',FALSE) ; SetParamSoc('SO_TAUXEURO',6.55957) ; SetParamSoc('SO_DECVALEUR',2) ;
  SetParamSoc('SO_DECPRIX',2) ;
  SetParamSoc('SO_DEVISEPRINC','FRF') ;
  END Else If LaMonnaie='E' Then
  BEGIN
  Q:=OPENSQL('SELECT * FROM DEVISE WHERE D_DEVISE="FRF" ',FALSE) ;
  If Q.Eof Then
    BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('D_DEVISE').AsString:='FRF' ;
    Q.FindField('D_LIBELLE').AsString:='Francs' ;
    Q.FindField('D_SYMBOLE').AsString:='F' ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.FindField('D_DECIMALE').AsInteger:=2 ;
    Q.FindField('D_QUOTITE').AsInteger:=1 ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END Else
    BEGIN
    Q.Edit ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END ;
  Ferme(Q) ;
  Q:=OPENSQL('SELECT * FROM DEVISE WHERE D_DEVISE="EUR" ',FALSE) ;
  If Q.Eof Then
    BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('D_DEVISE').AsString:='EUR' ;
    Q.FindField('D_LIBELLE').AsString:='Euro' ;
    Q.FindField('D_SYMBOLE').AsString:='EUR' ;
    Q.FindField('D_CODEISO').AsString:='EUR' ;
    Q.FindField('D_DECIMALE').AsInteger:=2 ;
    Q.FindField('D_QUOTITE').AsFloat:=1 ;
    Q.FindField('D_MONNAIEIN').AsString:='-' ;
    Q.FindField('D_FONGIBLE').AsString:='-' ;
    Q.FindField('D_PARITEEURO').AsFloat:=1 ;
    Q.FindField('D_FERME').AsString:='-' ;
    If VH^.RecupSISCOPGI Then
      BEGIN
      Q.FindField('D_MAXDEBIT').AsFloat:=1 ;
      Q.FindField('D_MAXCREDIT').AsFloat:=1 ;
      END ;
    Q.Post ;
    END Else
    BEGIN
    Q.Edit ;
    Q.FindField('D_MONNAIEIN').AsString:='-' ;
    Q.FindField('D_FONGIBLE').AsString:='-' ;
    Q.FindField('D_PARITEEURO').AsFloat:=1 ;
    Q.FindField('D_FERME').AsString:='-' ;
    If VH^.RecupSISCOPGI Then
      BEGIN
      Q.FindField('D_MAXDEBIT').AsFloat:=1 ;
      Q.FindField('D_MAXCREDIT').AsFloat:=1 ;
      END ;
    Q.Post ;
    END ;
  Ferme(Q) ;
  SetParamSoc('SO_DEVISEPRINC','EUR') ;
  SetParamSoc('SO_TENUEEURO',TRUE) ; SetParamSoc('SO_TAUXEURO',1) ; SetParamSoc('SO_DECVALEUR',2) ;
  SetParamSoc('SO_DECPRIX',2) ;
  END ;
END ;
{$ENDIF}

Type TFF = Array[1..10] Of TFourCpt ;

Procedure Dechiffre(St : String ; Var C : TFF) ;
Var St1,St2 : String ;
    SDeb,SFin : String ;
    i : Integer ;
    ATraiter : Boolean ;
    Tous : Boolean ;
BEGIN
i:=Pos(':',St) ;
If i>0 Then BEGIN St1:=Copy(St,1,i-1) ; St2:=Copy(St,i+1,Length(St)-i) END Else
  BEGIN St1:=St ; St2:=St ; END ;
For i:=1 To 10 Do
  BEGIN
  SDeb:=ReadTokenstV(St1) ; SFin:=ReadTokenstV(St2) ;
  ATRaiter:=(Pos('-',SDeb)<=0) And (Pos('#',SDeb)<=0) ; Tous:=Pos('*',SDeb)>0 ;
  If ATraiter Then
    BEGIN
    If Tous Then BEGIN C[i].Deb:='*' ; C[i].Fin:='*' ; End
            Else BEGIN C[i].Deb:=SDeb ; C[i].Fin:=SFin ; END ;
    END ;
  END ;
END ;

Function CompareTL(St1,St2 : String) : Boolean ;
Var C1,C2 : TFF ;
    i : Integer ;
    Tous1,Tous2,OkOk : Boolean ;
BEGIN
Result:=TRUE ;
Fillchar(C1,SizeOf(C1),#0) ; Fillchar(C2,SizeOf(C2),#0) ;
Dechiffre(St1,C1) ; Dechiffre(St2,C2) ;
For i:=1 To 10 Do
  BEGIN
  If C1[i].Deb='' Then Continue ; If C2[i].Deb='' Then Continue ;
  Tous1:=C1[i].Deb='*' ; Tous2:=C2[i].Deb='*' ;
  If Tous1 And (C2[i].Deb<>'') Then Exit Else If Tous2 And (C1[i].Deb<>'') Then Exit Else
    BEGIN
    If (Pos('###',C1[i].Deb)>0) Or (Pos('---',C1[i].Deb)>0) Or
       (Pos('###',C2[i].Deb)>0) Or (Pos('---',C2[i].Deb)>0) Then Continue ;
    OkOk:=((C1[i].Deb<C2[i].Deb) And (C1[i].Fin<C2[i].Deb)) Or
          ((C2[i].Deb<C1[i].Deb) And (C2[i].Fin<C1[i].Deb)) ;
    If Not OkOk THen Exit ;
    END ;
  END ;
Result:=FALSE ;
END ;

Function OnCpt(TOBL :TOB ; Compte,Sauf : String) : Boolean;
Var CPT : String ;
    Cpt1 : String ;
    ll : Integer ;
BEGIN
Result:=FALSE ;
Cpt:=TOBL.GetValue('CPT') ;
Compte:=Trim(Compte) ; Sauf:=Trim(Sauf) ;
If Compte='' Then BEGIN Result:=TRUE ; Exit ; END ;
If (Compte<>'') And (Compte[Length(Compte)]<>';') Then Compte:=Compte+';' ;
While Compte<>'' Do
  BEGIN
  Cpt1:=ReadTokenSt(Compte) ; ll:=Length(Cpt) ; If Length(Cpt1)<ll Then ll:=Length(Cpt1) ;
  If (Cpt1='TOUS') Or (Copy(Cpt,1,ll)=Copy(Cpt1,1,ll)) Then BEGIN Result:=TRUE ; Break ; END ;
  END ;
If Result And (Sauf<>'') Then
  BEGIN
  If (Sauf[Length(Sauf)]<>';') Then Sauf:=Sauf+';' ;
  While Sauf<>'' Do
    BEGIN
    Cpt1:=ReadTokenSt(Sauf) ; ll:=Length(Cpt) ; If Length(Cpt1)<ll Then ll:=Length(Cpt1) ;
    If Copy(Cpt,1,ll)=Copy(Cpt1,1,ll) Then BEGIN Result:=FALSE ; Break ; END ;
    END ;
  END ;
END ;

Function OnTL(TOBL :TOB ; TL : String) : Boolean ;
Var TL1,TLi : String ;
    i : Integer ;
BEGIN
If Trim(TL)='' Then BEGIN Result:=TRUE ; Exit ; END ; 
Result:=FALSE ; TL1:='' ;
For i:=0 To 9 Do
  BEGIN
  TLi:=Trim(TOBL.GetValue('TL'+IntToStr(i))) ;
  If TLi='' Then TL1:=TL1+'#,' Else TL1:=TL1+Tli+',' ;
  END ;
Delete(TL1,Length(TL1),1) ;
If CompareTL(TL1,TL) Then Result:=TRUE ;
END ;

{$IFNDEF NOVH}
Function EstDansCpt(OnGen : Boolean ; TOBL : TOB ; Ind : tProfilTraitement) : Boolean ;
Var Compte,Sauf,TL : String ;
    OkCpt,OkTL : Boolean ;
BEGIN
Result:=FALSE ; OkCpt:=FALSE ; OkTL:=FALSE ;
If OnGen Then
  BEGIN
  Compte:=VH^.ProfilUserC[Ind].Compte1 ;
  Sauf:=VH^.ProfilUserC[Ind].Exclusion1 ;
  TL:=VH^.ProfilUserC[Ind].TableLibre1 ;
  END Else
  BEGIN
  Compte:=VH^.ProfilUserC[Ind].Compte2 ;
  Sauf:=VH^.ProfilUserC[Ind].Exclusion2 ;
  TL:=VH^.ProfilUserC[Ind].TableLibre2 ;
  END ;
If OnCpt(TOBL,Compte,Sauf) Then BEGIN OkCpt:=TRUE ; END ;
If OnTL(TOBL,TL) Then BEGIN OkTL:=TRUE ; END ;
If OkCpt And OkTL Then Result:=TRUE ;
END ;

Function CptDansProfil(Gen,Aux : String ; Ind : tProfilTraitement) : Boolean ;
Var St : String  ;
    TC,TOBL : Tob ;
    Q : TQuery ;
    OkOnGen : Boolean ;
    OkOnAux : Boolean ;
BEGIN
//Result:=FALSE ;
If (Ind in [prClient,prFournisseur])=FALSE Then BEGIN Result:=TRUE ; Exit ; END ;
If Not VH^.ProfilCCMPExiste[Ind] Then BEGIN Result:=TRUE ; Exit ; END ;
OkOnGen:=TRUE ; OkOnAux:=TRUE ;
If Gen<>'' Then
  BEGIN
  St:='Select G_GENERAL AS CPT, G_TABLE0 AS TL0 , G_TABLE1 AS TL1, G_TABLE2 AS TL2, G_TABLE3 AS TL3,'+
      'G_TABLE4 AS TL4 , G_TABLE5 AS TL5, G_TABLE6 AS TL6, G_TABLE7 AS TL7, G_TABLE8 AS TL8, G_TABLE9 AS TL9 FROM GENERAUX WHERE G_GENERAL="'+Gen+'" ' ;
  TC:=TOB.Create('',Nil,-1) ;
  Q:=OpenSQL(St,TRUE,-1,'',true);
  TC.LoadDetailDB('CPT','','',Q,False,True) ;
  Ferme(Q) ;
  If TC.Detail.Count=1 Then
    BEGIN
    TOBL:=TC.Detail[0] ; OkOnGen:=EstDansCpt(TRUE,TOBL,Ind) ;
    END ;
  TC.Free ;
  END ;
If Aux<>'' Then
  BEGIN
  St:='Select T_AUXILIAIRE AS CPT, T_TABLE0 AS TL0 , T_TABLE1 AS TL1, T_TABLE2 AS TL2, T_TABLE3 AS TL3,'+
      'T_TABLE4 AS TL4 , T_TABLE5 AS TL5, T_TABLE6 AS TL6, T_TABLE7 AS TL7, T_TABLE8 AS TL8, T_TABLE9 AS TL9 FROM TIERS WHERE T_AUXILIAIRE="'+Aux+'" ' ;
  TC:=TOB.Create('',Nil,-1) ;
  Q:=OpenSQL(St,TRUE,-1,'',true);
  TC.LoadDetailDB('CPT','','',Q,False,True) ;
  Ferme(Q) ;
  If TC.Detail.Count=1 Then
    BEGIN
    TOBL:=TC.Detail[0] ; OkOnAux:=EstDansCpt(FALSE,TOBL,Ind) ;
    END ;
  TC.Free ;
  END ;
Result:=OkOnGen And OkOnAux ;
If Not Result Then
  BEGIN
  St:='0;ATTENTION ! Problème lié au profil;Vous n''avez pas les droit d''accès sur ce compte ('+Gen+' '+Aux+');E;O;O;O;' ;
  {$IFNDEF EAGLSERVER}
  HShowMessage(St,'','') ;
  {$ENDIF}
  END ;
END ;

Function MonProfilOk(Qui : tProfilTraitement) : Boolean ;
BEGIN
Result:=FALSE ;
{$IFNDEF CCMP}
Exit ;
{$ENDIF}
If Not VH^.ProfilCCMPExiste[Qui] Then Exit ;
If (qui in [prClient,prFournisseur])=FALSE Then Exit ;
If Not VH^.ProfilUserC[Qui].ProfilOk Then Exit ;
Result:=TRUE ;
END ;
{$ENDIF}

Function _ChampExiste(NomTable,NomChamp : String) : Boolean ;
BEGIN
{$IFDEF EAGLCLIENT}
  try
    Result := True;
    ExecuteSQL('SELECT '+NomChamp+' FROM '+NomTable);
  except
    on E: Exception do begin Result := False; Exit; end;
  end;
{$ELSE}
  Result:=ChampPhysiqueExiste(NomTable,NomChamp) ;
{$ENDIF}
END ;

{$IFDEF EURO}
Procedure EcritProcEuro(St : String) ;
Var FF : TextFile ;
    LaDir,NomFic,St1 : String ;
    io : Integer ;
BEGIN
{$i-}
LaDir:=ExtractFilePath(Application.ExeName) ;
NomFic:=LaDir+'EUROLOG.TXT' ;
AssignFile(FF,NomFic) ;
IF Not FileExists(NomFic) Then Rewrite(FF) Else Append(FF) ;
St1:='  '+V_PGI.User+' Le '+DateToStr(Date)+' : '+St ;
Writeln(FF,St1) ;
CloseFile(FF) ;
{$i+}
END ;
{$ENDIF}

{Function OkSynchro : Boolean ;
BEGIN
(*Result:=FALSE ;
If (ctxPCL in V_PGI.PGIContexte) Then Result:=TRUE ; //Else Result:=TRUE ;*)
Result := True; // Synchro PGE (Sx)
END ; }

FUNCTION GoINSERE( st,st1 : string ; deb,long : integer) : string ;
BEGIN
Result:=copy(st,1,deb-1)+copy(st1,1,Long)+copy(st,deb+long,length(st)-long) ;
END ;

{$IFNDEF NOVH}
Function HalteAuFeu : Boolean ;
BEGIN
Result:=FALSE ;
If VH^.StopRSP Then Result:=TRUE ;
END ;

Procedure DemandeStop(St : String) ;
BEGIN
If Not VH^.STOPRSP Then
  BEGIN
  {$IFNDEF EAGLSERVER}
  If HShowMessage('22;?Caption?;Confirmez-vous l''arrêt du traitement en cours ?;Q;YN;N;N;',St,'')=mrYes Then BEGIN VH^.STOPRSP:=TRUE ; Application.ProcessMessages ; End ;
  {$ENDIF}
  END ;
END ;
{$ENDIF}

Function EstSpecif(St : String) : Boolean ;
BEGIN
Result:=(Pos(St,V_PGI.Specif)>0) ;
//RRO 30012003 If St='51188' Then Result := Not Result;
{If Result Then
  BEGIN
  If St='51188' Then VH^.FinEuro:=TRUE ;
  END ;}
END ;


function EstMonnaieIN ( CodeD : String3 ) : boolean ;
BEGIN
{$IFNDEF NOVH}
Result:=Pos(';'+CodeD+';',VH^.MonnaiesIn)>0 ;
{$ELSE}
Result := TRUE;
{$ENDIF}
END ;

Function StopEcart (DD : tDateTime ; Dev : String = '') : Boolean ;
BEGIN
Result:=True ;
{If Not VH^.FinEuro Then Exit ;
//If dev='' Then
Result:=(dd>=encodeDate(2002,01,01)) ;}
END ;


Function EstbaseConso : Boolean ;
BEGIN
Result:=FALSE ;
{$IFDEF CONSOCERIC}
If VH^.IsBaseConso Then Result:=TRUE ;
{$ENDIF}
END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 28/03/2002
Modifié le ... : 17/09/2002
Description .. : retourne le IBAN d'un compte
Mots clefs ... : IBAN;COMPTE
*****************************************************************}
// Prévenir IMPERATIVEMENT la trésorerie lors de la modification de cette fonction
function calcIBAN(pays : string ; RIB : string) : string ;
Var St2, St4, cleIBAN, ret, strInter : String ;
		ii : Byte ;
		cleL : integer ;
		i : Integer ;
BEGIN
	//lRIB=litnombreIBAN(Banque+guichet+cle);
	//lPays=litnombreIBAN(pays);
	Result:='' ;
	St2:=trim(RIB)+Trim(Pays)+'00' ;
	if Length(St2)<10 then exit ;
	St2:=UpperCase(St2) ;
	//Transforme les lettres en chiffres selon le NEE 5.3
	i:=1 ;
	while i<Length(St2) do
	begin
		if St2[i] in ['A'..'Z'] then
		BEGIN
			ii:=Ord(St2[i])-65 ;
			st4:= copy(st2,1,i-1) + inttostr(10+ii) + copy(st2,i+1, length(st2));
			st2:=st4 ;
		END ;
		inc(i);
	end ;

	ret := '' ;
	cleL := 0 ;
	st4:='';
	//On découpe par tranche de 9
	//On calcul la clé via mod 97 puis on fait clé + reste du rib
        if QuelPaysLocalisation=CodeISOES then
           Begin //XMG 14/07/03 début (depuis le deuxième passage on prennait CleL+9 caractères de St2, ce qui n'est pas ce qu'on attendaît)....
             While length(st2)>0 do
             begin
              st2:= FormatFloat('####',cleL)+St2 ;
              st4 := copy(st2,1,9) ;
              delete(st2,1,9);
              cleL := strtoint64(st4) mod 97 ;
             End ;
           End else
           Begin
            for i:=1 to (length(st2) div 9)+1 do
            begin
                    st4 := copy(st2,1,9) ;
                    delete(st2,1,9);
                    strInter := inttostr(cleL)+st4 ;
                    cleL := strtoint64(strinter) mod 97 ;
            end ;
        End ; //XMG 24/02/2005
	//une fois fini, on calcul 98-clé
	cleIBAN := inttostr(98-(cleL  mod 97));
	if length(cleIBAN)=1 then  cleIBAN := '0' + cleIBAN ;
	Result:= trim(Pays)+ trim(cleIBAN)+ trim(RIB) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 17/09/2002
Modifié le ... :   /  /
Description .. : Mets en forme les RIB
Mots clefs ... :
*****************************************************************}
// Prévenir IMPERATIVEMENT la trésorerie lors de la modification de cette fonction
function calcRIB(pays : string ; banque : string ; guichet : string ; compte : string; cle : string) : string ;
begin
{$IFDEF COMPTA}
   if QuelPaysLocalisation=CodeISOES then
      Result:=FindetReplace(FindetReplace(EncodeRib(Banque,Guichet,Compte,Cle,'',Pays),' ','',TRUE),'/','',TRUE)
   else Begin
{$ENDIF}
     if (pays='ES') then
       result := trim(banque) + trim(guichet) + trim(cle) + trim(compte)
     else if (pays='FR') then
       result := trim(banque) + trim(guichet) + trim(compte) + trim(cle)
     else
       result := trim(banque) + trim(guichet) + trim(compte) + trim(cle);
{$IFDEF COMPTA}
   End ; //XVI 24/02/2005
{$ENDIF}
//XMG 14/07/03 fin
end ;

//----------------------------------------------------
//--- Nom   : IsDossierCabinet
//--- Objet : Détermine si c'est un dossier cabinet
//----------------------------------------------------
function IsDossierCabinet (NumDossier : String) : Boolean;
var ChSql      : String;
    RSql       : Tquery;
    ChResultat : Boolean;
begin
 ChResultat:=False;
 ChSql:='Select DOS_CABINET From DOSSIER where DOS_NODOSSIER="'+NumDossier+'"';
 RSql:=OpenSql (ChSql,True,-1,'',true);
 if (Not RSql.Eof) then
  ChResultat:=RSql.FindField ('DOS_CABINET').AsString='X';
 Ferme (RSql);

 IsDossierCabinet:=ChResultat;
end;

{$IFDEF COMPTA}

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 14/10/2003
Modifié le ... :   /  /    
Description .. : function de recopie des groupe d'utilisateurs S7 vers S5 
Suite ........ : pack avancé
Mots clefs ... : 
*****************************************************************}
procedure ExportS7ImportS5Confidentialite;
var Q : TQuery;
    Corresp,i : integer ;
    AFaire, DroitS7 : boolean ;
    stSQL, stMn1, stAccesGrp, stTag,s : string ;
    l : TStringList ;
BEGIN
   //On considère que s'il existe du paramétrage sur S7, on peut recopié
   DroitS7 := ExisteSQL('select mn_accesgrp  from menu where mn_1=1 and mn_accesgrp like "%-%"');
   if (V_PGI.FSuperviseur and DroitS7) then
   begin
      AFaire := not GetParamSocSecur('SO_CONFIDCOPIE',False) ;
      if AFaire then // attention en cours de test
      begin
         l := TStringList.create ;
         //Récupération des accès menu 7
         Q := OpenSQL('SELECT MN_TAG,MN_ACCESGRP,MN_1 FROM MENU WHERE MN_1 IN (1,2,3,6,7,15,25,26) ORDER BY MN_TAG', True,-1,'',true);
         if not Q.EOF then
            InitMove(Q.RecordCount-1,'Récupération');
         while not Q.EOF do
         begin
            movecur(false);
            if Q.FindField('MN_TAG').AsInteger <> 0 then
            begin
               s := Q.FindField('MN_ACCESGRP').AsString;
               s := IntToStr(Q.FindField('MN_TAG').AsInteger) + ';' + IntToStr(Q.FindField('MN_1').AsInteger) + ';' + s;
               l.Add(s);
            end;
            Q.Next;
         end;
         FiniMove;
         InitMove(l.count-1,'Recopie');
         for i:=0 to l.Count-1 do
         begin
             //sauvegarde dans accés menu de S5
            s := l[i];
            stTag := ReadTokenSt(s); //t=tag, il reste mn1 et l'accès dans s
            stMn1 := ReadTokenSt(s); //t=mn_1, il reste l'accès dans s
            stAccesGrp := s;
            corresp := strtoint(stTag) ;
            //Correspondance entre S7 et S5
            case corresp of
               7106 : corresp:=7112;
               7109 : corresp:=26012;
               7121 : corresp:=26011;
               7139 : corresp:=7145;
               7142 : corresp:=26013;
               7145 : corresp:=26014;
               7172 : corresp:=7178;
               7175 : corresp:=26015;
               7178 : corresp:=26016;
               7205 : corresp:=7211;
               7208 : corresp:=26017;
               7211 : corresp:=26018;
               7242 : corresp:=7005;
               7781 : corresp:=7787;
               7784 : corresp:=7787;
               1430 : corresp:=7352;
               1435 : corresp:=7355;
            end ;
            movecur(false);
            stSQL := 'UPDATE MENU SET MN_ACCESGRP="' ;
            stSQL := stSQL + stAccesGrp ;
            stSQL := stSQL + '" WHERE MN_TAG=' ;
            stSQL := stSQL + inttostr(Corresp) ;
            //stSQL := stSQL + ' AND MN_1 in (8,9,11,12,13,14,16,17,18,20)' ;
            stSQL := stSQL + ' AND MN_1 in (6,8,9,11,12,13,14,16,17,18,20,27)';
            BeginTrans;
            try
               ExecuteSQL(stSQL);
               CommitTrans ;
            except
               V_PGI.IoError:=oeUnknown;
               Rollback;
            end;
         end ; // For i
         if V_PGI.IoError<>oeUnknown then
         begin
            stSQL := 'UPDATE MENU SET MN_ACCESGRP="0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" WHERE MN_1=0 AND MN_2 in (6,8,9,11,12,13,14,16,17,18,20,27)';
            begintrans;
            try
               executeSQL(stSQL);
               commitTrans;
               SetParamSoc('SO_CONFIDCOPIE',true);
            except
               SetParamSoc('SO_CONFIDCOPIE',False);
               RollBack;
            end ;
         end
         else
         begin
            SetParamSoc('SO_CONFIDCOPIE',False);
         end ;
         FreeAndNil(l);//on libère la ram
      end ; //IF AFaire
      FiniMove;
   end ;//packavance
END ;

{$ENDIF EAGLSERVER}
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/10/2003
Modifié le ... :   /  /
Description .. : Force la propriété V_PGI.CascadeForms à vrai dans le cas
Suite ........ : PCL.
Suite ........ : Pour activer cette fonctionnalité, il faut une résolution
Suite ........ : supérieure à 800x600
Mots clefs ... :
*****************************************************************}
procedure ForceCascadeForms;
begin
  { CascadeForms est forcé uniquement en PCL }
{$IFNDEF HVCL}
{$IFNDEF EAGLSERVER}
  if V_PGI.ModePCL='1' then
  begin
    if ((Screen.Width>800) and (Screen.Height>600)) then
    begin
      // Si CascadeForms n'a jamais été renseigné, on obtient 2 en retour ( valeur par défaut )
      if GetSynRegKey('CascadeForms',2,True)=2 then
        SaveSynRegKey('CascadeForms',True,True);
    end;
  end;
{$ENDIF EAGLSERVER}
{$ENDIF HVCL}
end;

{Fonction un peu plus exigeante que IsValideDate et qui teste si la date en
 paramêtre est comprise entre 01/01/1900 et 31/12/2099}
function CIsValideDate(aDate : string) : Boolean;
begin
  Result := IsValidDate(aDate);
  if Result then
    Result := (iDate1900 <= StrToDate(aDate)) and (iDate2099 >= StrToDate(aDate));
end;

Function EstUneLigneCpt(St : String) : Boolean ;
Var Cod : String ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','EstUneLigneCpt') ;
{$ENDIF}
Cod:=Copy(St,1,3) ;
Result:=(Cod='¥¥¥') Or (Cod='***') Or (Cod='###') ;
END ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/07/2004
Modifié le ... :   /  /
Description .. : Détermine si la configuration du dossier autorise les modfis
Suite ........ : de pointage pour savoir qui a raison l'expert ou le client
Suite ........ : Valeur de retour : True si en mode Consultation
Suite ........ :                    False si en mode Création/Modification
Mots clefs ... :
*****************************************************************}
function CEstPointageEnConsultationSurDossier : boolean;
begin
  if CtxPcl in V_PGI.PGIContexte then
  begin
    Result :=  (GetParamsoc ('SO_CPPOINTAGESX') = 'CLI') and (((GetParamsoc ('SO_CPLIENGAMME') = 'S1')
      and (GetParamsoc ('SO_CPMODESYNCHRO')=True)) or (GetParamsoc ('SO_CPLIENGAMME') = 'S5'));
  end
  else
  begin // Contexte PGE
    if (GetParamsoc ('SO_CPLIENGAMME') = 'AUC') or (GetParamsoc ('SO_CPLIENGAMME') = 'SI') or (GetParamsoc ('SO_CPLIENGAMME') = '') then
      Result := False
    else
      Result := (GetParamsoc ('SO_CPPOINTAGESX') = 'EXP');
  end;
end;

{$IFNDEF EAGLSERVER}

function _IsDossierNetExpert : Boolean;
var Q : TQuery;
begin
Result := False;
      Q := OpenSql ('Select DOS_NETEXPERT,DOS_NECPSEQ From DOSSIER where DOS_NODOSSIER="'+V_PGI.NoDossier+'"',True,-1,'',true);
      if not Q.Eof then
        Result := Q.FindField ('DOS_NETEXPERT').AsString='X';
      Ferme (Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 05/01/2005
Modifié le ... : 19/09/2007
Description .. : Renvoie les branches du paramsoc à virer du paramétrage
Suite ........ : société.
Suite ........ : CA - 19/09/2007 - Utilisation de cette fonction pour CCMP.
Mots clefs ... : 
*****************************************************************}
function BrancheParamSocAVirer : string;
begin
  Result := 'SCO_CPEDI';  // Pour l'instant - cf FQ 16076 - CA - 31/08/2005
  if V_PGI.FSuperviseur then
    Result := Result
  else
    if _IsDossierNetExpert then Result := Result + ';SCO_CPREVISION';

  Result := Result + ';SCO_OLAP';
  {$IFDEF CCMP}
  Result := Result + ';SCO_ICC';
  {$ENDIF}
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/10/2004
Modifié le ... :   /  /
Description .. : Renvoi la condition SQL qiu gère la confidentialité
Mots clefs ... :
*****************************************************************}
function CGenereSQLConfidentiel( vStPrefixe : string ) : string;
var lSt : string;
begin
  lSt := '';
  if (V_Pgi.Confidentiel = '0') then
  begin
    lSt := lSt + '((' + vStPrefixe + '_CONFIDENTIEL = "0") OR ' +
                 '(' + vstPrefixe + '_CONFIDENTIEL = "-")) ';
  end
  else
  begin
    lSt := lSt + '(((' + vStPrefixe + '_CONFIDENTIEL = "-") OR ' +
                 '(' + vStPrefixe + '_CONFIDENTIEL = "X")) OR ' +
                 '(' + vStPrefixe + '_CONFIDENTIEL <= "' + V_PGI.Confidentiel + '"))';
  end;
  Result := lSt;
end;

////////////////////////////////////////////////////////////////////////////////

Function VERIFNIF_ES ( NIF : String ) : String ;
Const LDNI : array[0..22] of char =
      ('T','R','W','A','G','M','Y','F','P','D','X','B','N','J','Z','S','Q','V','H','L','C','K','E') ;
Var
   Digit,First,CodeISO,QuelesNum : String ;
   Calcul,ii,pair,impair,jj      : integer ;
   DNI                           : Longint ;
   IsCIF                         : Boolean ;
   Interdits                     : set of char ;
BEGIN
  NIF:=uppercase(Trim(NIF)) ;
  Result:=NIF ;
  CodeISO:='' ;
  if copy(NIF,1,length(CodeISOES))=CodeISOES then begin
     CodeISO:=Copy(NIF,1,length(CodeISOES)) ;
     Delete(NIF,1,length(CodeISOES)) ;
  end ;
  IsCIF:=(pos(';'+Copy(NIF,1,1)+';',';A;B;C;D;E;F;G;P;Q;S;')>0) ;
  ii:=Length(NIF) ;
  interdits:=['/','\',',','.',';','-',' '] ;
  //DigitOld:=Copy(NIF,length(NIF),1) ;
  Delete(NIF,length(NIF),1) ; //on supprime le Digit de controle Ancien....
  First:='' ;
  if Not IsCIF then begin
     if Copy(NIF,1,1)='X' then begin       //C'est le NIF d'un étranger....
        First:='X' ;
        Delete(NIF,1,1) ;
     End ;
     Interdits:=Interdits+['A'..'Z'] ;
  end
  else begin
     First:=Copy(NIF,1,1) ;
     delete(NIF,1,1) ;
  end ;
  (*While ii>0 do begin
     if (NIF[ii] in Interdits) {and
        ((not isCIF) or (ii>1) and (NIF[ii]='X'))} then
        Delete(NIF,ii,1) ;
     Dec(ii) ;}
  End ;*)
  if IsCIF then begin //C'est un CIF
     Pair:=0 ; impair:=0 ; jj:=0 ;
     for ii:=1 to length(NIF) do begin
         if not (NIF[ii] in Interdits) then begin
            inc(jj) ;
            if (jj mod 2)=0 then
               Pair:=Pair+valeuri(Copy(NIF,ii,1))
            else Begin
               Calcul:=valeuri(Copy(NIF,ii,1))*2 ;
               Impair:=Impair+(Calcul div 10)+(Calcul mod 10) ;
            end ;
         End ;
     end ;
     Calcul:=Pair+Impair ; Calcul:=(((Calcul div 10)+1)*10)-Calcul ;
     if pos(';'+First+';',';Q;P;S;')>0 then
        Digit:=chr(64+Calcul)
     else
        Digit:=chr(48+(calcul mod 10)) ;
  end
  else Begin //C'est un DNI/NIF
     QuelesNum:='' ;
     For ii:=1 to length(NIF) do
         if not (NIF[ii] in Interdits) then
            QuelesNum:=QuelesNum+NIF[ii] ;
     Val(Copy(QuelesNum,1,length(QueLesNum)),DNI,ii) ;
     if ii=0 then begin
        Calcul:=(Round(DNI) mod 23) ;
        Digit:=LDNI[Calcul] ;
     End ;
  End;
  Result:=CodeISO+First+NIF+Digit ;

END ;


function OkChar( St : String ) : Boolean ;
var i : Integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','OkChar') ;
{$ENDIF}
 Result := false ;
 for i:=1 To Length(St) do if not (St[i] in ['0'..'9']) then exit ;
 Result := true ;
end ;

constructor TZCatBud.Create(vDossier : string = '');
Var fb : TFichierBase ;
    Q : TQuery ;
    StCode,StLib,StLibre,StS,StJ,St : String ;
    i,k,l : Integer ;
begin
  {$IFDEF TTW}
   cWA.MessagesAuClient('COMSX.IMPORT','','TZCatBub.Create') ;
  {$ENDIF}
  Fillchar(FCatBud, SizeOf(FCatBud), #0);
  Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="CJB"', TRUE, -1, 'CHOIXCOD', true , vDossier);
  k := 1;
  while (not Q.Eof) and (k <= MaxCatBud) do begin
    StCode := Q.FindField('CC_CODE'   ).AsString;
    StLib  := Q.FindField('CC_LIBELLE').AsString;
    StLibre:= Q.FindField('CC_LIBRE'  ).AsString;
    fb     := AxeTofb(Q.FindField('CC_ABREGE').AsString);
    i := Pos('/', StLibre);

    if i > 0 then begin
      StS := Copy(StLibre, 1, i - 1);
      StJ := Copy(StLibre, i + 1, Length(StLibre) - i);
      FCatBud[k].fb   := fb;
      FCatBud[k].Code := StCode;
      FCatBud[k].Lib  :=StLib;

      l := 1;
      while StS <> '' do begin
        St := ReadTokenSt(StS);
        if St <> '' then begin
          FCatBud[k].SurSect[l] := St;
          Inc(l);
        end;
      end;

      l := 1;
      while StJ <> '' do begin
       St := ReadTokenSt(StJ);
       if St <> '' then begin
         FCatBud[k].SurJal[l] := St;
         Inc(l);
       end;
     end;
     Inc(k);
    end;
    Q.Next;
  end;
  Ferme(Q);
end;

function TZCatBud.GetCatBud(Value : Integer) : TUneCatBud;
begin
  {$IFDEF TTW}
   cWA.MessagesAuClient('COMSX.IMPORT','','TZCatBub.GetCatBud') ;
  {$ENDIF}
   Result := FCatBud[Value];
end;

constructor TZInfoCpta.Create ( vDossier : String = '' ) ;
var
 StB            : string ;
 QLoc , QLoc1,Q : TQuery ;
 NumAxe         : integer ;
 fb             : TFichierBase ;
 CO,CL          : String ;
 Lg,i,j         : integer;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.Create') ;
{$ENDIF}
 MPACC := TStringList.Create;

 FDossier := vDossier ;

 ChargeLgDossier ;

 QLoc:=OpenSQL('Select * From STRUCRSE Where SS_CONTROLE="X" ',True,-1, 'STRUCRSE',true , FDossier) ;
 While Not QLoc.Eof do
   BEGIN
   fb:=AxeToFb(QLoc.FindField('SS_AXE').AsString) ;
   FCpta[fb].UneStruc.CodeSp:=QLoc.FindField('SS_SOUSSECTION').AsString ;
   FCpta[fb].UneStruc.Debu:=QLoc.FindField('SS_DEBUT').AsInteger ;
   FCpta[fb].UneStruc.Long:=QLoc.FindField('SS_LONGUEUR').AsInteger ;
   QLoc1:=OpenSql('Select PS_CODE From SSSTRUCR Where PS_AXE="'+QLoc.FindField('SS_AXE').AsString+'" And '+
                  'PS_SOUSSECTION="'+QLoc.FindField('SS_SOUSSECTION').AsString+'" ',True,-1,'',true,FDossier) ;
   FCpta[fb].UneStruc.QuelCarac:=QLoc1.Fields[0].AsString ; Ferme(QLoc1) ;
   QLoc.Next ;
   END ;
 Ferme(QLoc) ;

 NumAxe:=0 ;
 Q:=OpenSQL('SELECT * FROM AXE ORDER BY X_AXE',TRUE,-1,'AXE',true,FDossier) ;
 While Not Q.EOF do
    BEGIN
    Inc(NumAxe) ; fb:=AxeToFb(Q.FindField('X_AXE').AsString) ;
    FCpta[fb].Lg:=Q.FindField('X_LONGSECTION').AsInteger ;
    Stb:=Q.FindField('X_BOURREANA').AsString ;
    if Stb<>'' then
       BEGIN
       FCpta[fb].Cb:=Stb[1] ;
       END else
       BEGIN
       FCpta[fb].Cb:='0' ;
       END ;
    FSaisieTranche[NumAxe]:=(Q.FindField('X_SAISIETRANCHE').AsString='X') ;
    FCpta[fb].Chantier:=Q.FindField('X_CHANTIER').AsString='X' ;
    FCpta[fb].Structure:=Q.FindField('X_STRUCTURE').AsString='X' ;
    FCpta[fb].Attente:=Q.FindField('X_SECTIONATTENTE').AsString ;
    FCpta[fb].AxGenAttente:=Q.FindField('X_GENEATTENTE').AsString ;
    Q.Next ;
    END ;
   Ferme(Q) ;

 // Tables libres
   Fillchar(FLgTableLibre, SizeOf(FLgTableLibre),#0) ;
   Q:=OpenSQL('SELECT CC_CODE,CC_LIBRE,CC_TYPE FROM CHOIXCOD WHERE CC_TYPE="NAT" ORDER BY CC_TYPE,CC_CODE',TRUE,-1,'CHOIXCOD',true,FDossier) ;
   While Not Q.EOF Do
     BEGIN
     CO:=Q.Fields[0].AsString ; CL:=Q.Fields[1].AsString ;
     If OkChar(CL) Then
        BEGIN
        Lg:=StrToInt(CL) ;
        Case CO[1] Of
             'G' : i:=1 ; 'T' : i:=2 ; 'S' : i:=3 ; 'B' : i:=4 ;
             'D' : i:=5 ; 'E' : i:=6 ; 'A' : i:=7 ; 'U' : i:=8 ;
             'I' : i:=9 ;
             Else i:=0 ;
           END ;
        If i>0 Then BEGIN j:=StrToInt(Copy(CO,2,2))+1 ; FLgTableLibre[i,j]:=Lg ; END ;
        END ;
     Q.Next ;
     END ;
   Ferme(Q) ;

 ChargeSousPlanAxe ;
 ChargeMPACC ;

end;

destructor TZInfoCpta.Destroy;
var
 fb : TFichierBase ;
 i  : integer ;
begin

 for fb := fbAxe1 to fbAxe5 do
  for i := 1 To MaxSousPlan do
   if FSousPlanAxe[fb,i].ListeSP <> nil then
    begin
      FSousPlanAxe[fb,i].ListeSP.Clear ;
      FSousPlanAxe[fb,i].ListeSP.Free ;
    end ;

  MPAcc.Free;
 inherited ;

end;

procedure TZInfoCpta.ChargeLgDossier ;
var
 Stb : string ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.ChargeLgDossier') ;
{$ENDIF}

 if local then
   begin
   FCpta[fbGene].Lg      := GetParamSocSecur('SO_LGCPTEGEN', 0 ) ;
   FCpta[fbGene].Attente := GetParamSocSecur('SO_GENATTEND', '') ;
   Stb                   := GetParamSocSecur('SO_BOURREGEN', '') ;
   end
 else
   begin
   FCpta[fbGene].Lg      := GetParamSocDossierSecur('SO_LGCPTEGEN', 0 , FDossier) ;
   FCpta[fbGene].Attente := GetParamSocDossierSecur('SO_GENATTEND', '', FDossier) ;
   Stb                   := GetParamSocDossierSecur('SO_BOURREGEN', '', FDossier) ;
   end ;

 if Stb <> '' then
  FCpta[fbGene].Cb := Stb[1]
   else
      FCpta[fbGene].Cb := '0' ;

 if local then
   begin
   FCpta[fbAux].Lg       := GetParamSocSecur('SO_LGCPTEAUX', 0) ;
   Stb                   := GetParamSocSecur('SO_BOURREAUX', '') ;
   end
 else
   begin
   FCpta[fbAux].Lg       := GetParamSocDossierSecur('SO_LGCPTEAUX', 0,  FDossier) ;
   Stb                   := GetParamSocDossierSecur('SO_BOURREAUX', '', FDossier) ;
   end ;

 if Stb <> '' then
  FCpta[fbAux].Cb := Stb[1]
   else
    FCpta[fbAux].Cb:='0' ;

end ;


procedure TZInfoCpta.ChargeSousPlanAxe ;
var
 fb                : TFichierBase ;
 Q,Q1              : TQuery ;
 QAxe,QCode,QLib   : String ;
 QDeb,QLong        : Integer ;
 TabI              : array[fbAxe1..fbAxe5] Of Integer ;
 i                 : integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.ChargeSousPlanAxe') ;
{$ENDIF}

 for fb := fbAxe1 to fbAxe5 do
  for i := 1 To MaxSousPlan do
   if FSousPlanAxe[fb,i].ListeSP <> nil then
    begin
      FSousPlanAxe[fb,i].ListeSP.Clear ;
      FSousPlanAxe[fb,i].ListeSP.Free ;
    end ;

 Fillchar(FSousPlanAxe,SizeOf(FSousPlanAxe),#0) ;
 Fillchar(TabI,SizeOf(TabI),#0) ;

 Q := OpenSQL('Select * From STRUCRSE Order By SS_AXE, SS_DEBUT',True,-1,'STRUCRSE',true,FDossier) ;
 while not Q.Eof Do
  begin
   QAxe   := Q.FindField('SS_AXE').AsString ;
   QCode  := Q.FindField('SS_SOUSSECTION').AsString ;
   QLib   := Q.FindField('SS_LIBELLE').AsString ;
   QDeb   := Q.FindField('SS_DEBUT').AsInteger ;
   QLong  := Q.FindField('SS_LONGUEUR').AsInteger ;
   fb     :=AxeTofb(QAxe) ;
   Inc(TabI[fb]) ;
  if tabI[fb] <= MaxSousPlan Then
   begin
    FSousPlanAxe[fb,TabI[fb]].Code       := QCode ;
    FSousPlanAxe[fb,TabI[fb]].Lib        := QLib ;
    FSousPlanAxe[fb,TabI[fb]].Debut      := QDeb ;
    FSousPlanAxe[fb,TabI[fb]].Longueur   := QLong ;
    FSousPlanAxe[fb,TabI[fb]].ListeSP    := HTStringList.Create ;
    Q1 := OpenSQL('Select * from SSSTRUCR WHERE PS_AXE="' + QAxe + '" AND PS_SOUSSECTION="' + QCode + '" ORDER BY PS_CODE' , true , -1 ,'SSSTRUCR',true,FDossier) ;
    while not Q1.Eof Do
     begin
      FSousPlanAxe[fb,TabI[fb]].ListeSP.Add(Q1.FindField('PS_CODE').ASString + ';' + Q1.FindField('PS_LIBELLE').ASString + ';' ) ;
      Q1.Next ;
     end ;
    Ferme(Q1) ;
   end ;
   Q.Next ;
  end ;
 Ferme(Q) ;
end ;

function TZInfoCpta.GetCpta(Value: TFichierBase) : TInfoCpta;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.GetCpta') ;
{$ENDIF}
 result := FCpta[Value];
end;

function TZInfoCpta.GetLgTableLibre ( i , j : Integer) : Integer;
begin
 result := FLgTableLibre[i, j];
end;

function TZInfoCpta.GetSousPlanAxe( fb : TFichierBase ; vNumPlan : integer ) : TSousPlan;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.GetSousPlanAxe') ;
{$ENDIF}
 result := FSousPlanAxe[fb,vNumPlan] ;
end;

{ TZHalleyUser }

constructor TZHalleyUser.Create ( vDossier : String = '' ) ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TZHalleyUser.Create') ;
{$ENDIF}

 FDossier := vDossier ;

 ChargeHalleyUser ;
end;

procedure TZHalleyUser.ChargeHalleyUser ;
var
 lQG : TQuery ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TZHalleyUser.ChargeHalleyUser') ;
{$ENDIF}

 FGrpMontantMin := 0 ;
 FGrpMontantMax :=999999999999.0 ;
 FJalAutorises  := '' ;

 lQG := OpenSQL('SELECT UG_NIVEAUACCES, UG_NUMERO, UG_LIBELLE, UG_CONFIDENTIEL, UG_MONTANTMIN, UG_MONTANTMAX, UG_JALAUTORISES, UG_LANGUE, UG_PERSO FROM USERGRP WHERE UG_GROUPE="' + V_PGI.Groupe + '"', TRUE , -1, 'USERGRP', true ,FDossier) ;

 if not lQG.EOF then
  begin

   FGrpMontantMin := lQG.FindField('UG_MONTANTMIN').AsFloat ;
   FGrpMontantMax := lQG.FindField('UG_MONTANTMAX').AsFloat ;
   FJalAutorises  := lQG.FindField('UG_JALAUTORISES').AsString ;

   // CA - 18/07/2007 - FQ 20027 - Cas du << Tous >> enregistré dans la base
   if (Pos('<<',FJalAutorises)>0) then FJalAutorises:= '';

   if FJalAutorises = '-' then FJalAutorises := '' ; // 13569

   if FJalAutorises <> '' then
    begin
      if FJalAutorises[1] <> ';' then FJalAutorises := ';' + FJalAutorises ;
      if FJalAutorises[Length(FJalAutorises)] <> ';' then FJalAutorises := FJalAutorises + ';' ;
    end ;

   if ( FGrpMontantMax <= FGrpMontantMin) then
    begin
     FGrpMontantMin := 0 ;
     FGrpMontantMax :=999999999999.0 ;
    end ;

 end ;

 Ferme(lQG) ;

end;


procedure CPStatutDossier ( vStDossier : string = '' ) ;
begin
{$IFNDEF EAGLSERVER}
{$IFNDEF NOVH}
  if ( VH^.BStatutLiasse <> nil ) and ( VH^.BStatutLiasse.ImageIndex <> 3 ) then
  begin
    SetParamSoc('SO_CPSTATUTDOSSIERLIASSE', 3);
    VH^.BStatutLiasse.ImageIndex := 3;
  end ; // if

  if ( VH^.BStatutRevision <> nil ) and ( VH^.BStatutRevision.ImageIndex <> 3 ) then
  begin
    SetParamSoc('SO_CPSTATUTREVISION', 3);
    VH^.BStatutRevision.ImageIndex := 3;
  end;
{$ENDIF}
{$ENDIF}
end ;


Function GetTenuEuro : Boolean;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetTenueEuro') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := GetParamsoc ('SO_TENUEEURO');
{$ELSE}
 Result := VH^.TenueEuro;
{$ENDIF}
end;

Function GetRecupPCL : Boolean;
begin
{$IFDEF TTW}
 //cWA.MessagesAuClient('COMSX.IMPORT','','GetRecupPCL') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := TRUE; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupPCL;
{$ENDIF}
end;

Function GetFromPCL : Boolean;
begin
{$IFDEF TTW}
 //cWA.MessagesAuClient('COMSX.IMPORT','','GetFromPCL') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := TRUE; // Pour La Comsx Server
{$ELSE}
 Result := VH^.FromPCL;
{$ENDIF}
end;

Function GetRecupComSx : Boolean;
begin
{$IFDEF TTW}
 //cWA.MessagesAuClient('COMSX.IMPORT','','GetRecupComSx') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := TRUE; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupComSx;
{$ENDIF}
end;

Function GetRecupSISCOPGI : Boolean;
begin
{$IFDEF TTW}
// cWA.MessagesAuClient('COMSX.IMPORT','','GetRecupSISCOPGI') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := FALSE; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupSISCOPGI;
{$ENDIF}
end;

Function GetRecupLTL : Boolean;
begin
{$IFDEF TTW}
// cWA.MessagesAuClient('COMSX.IMPORT','','GetRecupLTL') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := FALSE; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupLTL;
{$ENDIF}
end;

Function GetRecupCegid : Boolean;
begin
{$IFDEF TTW}
 //cWA.MessagesAuClient('COMSX.IMPORT','','GetRecupCegid') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := FALSE; // Pour La Comsx Server
{$ELSE}
 Result := VH^.RecupCegid;
{$ENDIF}
end;

Function GetCPIFDEFCEGID : Boolean;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetCPIFDEFCEGID') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := GetParamSocSecur('SO_IFDEFCEGID',False); // Pour La Comsx Server
{$ELSE}
 Result := VH^.CPIFDEFCEGID;
{$ENDIF}
end;


function GetInfoCpta ( Value : TFichierBase ) : TInfoCpta ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetInfoCpta') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.InfoCpta.Cpta[Value] ;
{$ELSE}
 Result := VH^.Cpta[Value];
{$ENDIF}
end ;

function GetInfoAcceptation : TStrings ;
begin
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.InfoCpta.MPACC ;
{$ELSE}
 Result := VH^.MPACC;
{$ENDIF}
end;

function GetLgTableLibre(I, J : integer) : integer;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetLgTableLibre') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.InfoCpta.LgTableLibre[I, J];
{$ELSE}
 result := VH^.LgTableLibre[i,j];
{$ENDIF}
end;

function GetSousPlanAxe(fb : TFichierBase ; vNumPlan : integer) : TSousPlan ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetSousPlanAxe') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.InfoCpta.SousPlanAxe[fb,vNumPlan] ;
{$ELSE}
 result := VH^.SousPlanAxe[fb,vNumPlan] ;
{$ENDIF}
end;

function GetCatBud(aCatBud : Integer) : TUneCatBud;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','GetCatBud') ;
  {$ENDIF}
  
  {$IFDEF NOVH}
  Result := TCPContexte.GetCurrent.InfoCatBud.CatBud[aCatBud];
  {$ELSE}
  Result := VH^.CatBud[aCatBud];
  {$ENDIF}
end;

function GetCPExoRef : TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetRecupPCL') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.Exercice.CPExoRef ;
{$ELSE}
  Result := VH^.CPExoRef;
{$ENDIF}
end ;

function GetExercices(n : Integer) : TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetExercice') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.Exercice.Exercices[n] ;
{$ELSE}
  Result := VH^.Exercices[n];
{$ENDIF}
end ;


function GetEntree : TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetRecupPCL') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.Exercice.Entree ;
{$ELSE}
  Result := VH^.Entree;
{$ENDIF}
end ;

function GetEnCours : TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetEnCours') ;
{$ENDIF}
{$IFDEF NOVH}
Result := TCPContexte.GetCurrent.Exercice.EnCours ;
{$ELSE}
  Result := VH^.Encours;
{$ENDIF}
end ;

function GetNMoins2 : TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetNMoins2') ;
{$ENDIF}
{$IFDEF NOVH}
Result := TCPContexte.GetCurrent.Exercice.NMoins2 ;
{$ENDIF}
end ;

function GetPrecedent: TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetPrecedent') ;
{$ENDIF}
{$IFDEF NOVH}
 result    := TCPContexte.GetCurrent.Exercice.Precedent ;
{$ELSE}
  Result := VH^.Precedent;
{$ENDIF}
end ;

function GetExoClo:  TTabExo ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetExoClo') ;
{$ENDIF}
{$IFNDEF NOVH}
Result := VH^.ExoClo;
{$ENDIF}
end ;

function GetSuivant : TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetSuivant') ;
{$ENDIF}
{$IFDEF NOVH};
  result    := TCPContexte.GetCurrent.Exercice.Suivant ;
{$ELSE}
  Result := VH^.Suivant;
{$ENDIF}
end ;

function GetGrpMontantMin : double ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetGrpMontantMin') ;
{$ENDIF}
{$IFDEF NOVH}
 result    := TCPContexte.GetCurrent.HalleyUser.GrpMontantMin ;
{$ELSE}
 Result := VH^.GrpMontantMin;
{$ENDIF}
end ;


function GetGrpMontantMax : double ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetGrpMontantMax') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.HalleyUser.GrpMontantMax ;
{$ELSE}
 Result := VH^.GrpMontantMax;
{$ENDIF}
end ;

function GetJalAutorises : string ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetJalAutorises') ;
{$ENDIF}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.HalleyUser.JalAutorises ;
{$ELSE}
 Result := VH^.JalAutorises;
{$ENDIF}
end ;



Function GetBourreLibre : Char;
{$IFDEF NOVH}
var
   BLibre : String;
{$ENDIF}
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetBourreLibre') ;
{$ENDIF}
{$IFDEF NOVH}
      Result := #0 ;
     BLibre := GetParamSocSecur('SO_BOURRELIB', TRUE);
     if BLibre<>'' then Result := BLibre[1] ;
{$ELSE}
     Result := VH^.BourreLibre;
{$ENDIF}
end;

Function GetMultisouche : Boolean;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetMultisouche') ;
{$ENDIF}
{$IFDEF NOVH}
Result := FALSE;
{$ELSE}
Result := VH^.MultiSouche ;
{$ENDIF}
end;

Function GetModCPPackAvance : Boolean;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetModCPPackAvance') ;
{$ENDIF}
{$IFDEF NOVH}
Result := FALSE;
{$ELSE}
Result := VH^.OkModCPPackAvance ;
{$ENDIF}
end;

Function GetAnaCroisaxe : Boolean;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetAnaCroisaxe') ;
{$ENDIF}
{$IFDEF NOVH}
Result := GetParamSocSecur('SO_CROISAXE',False);
{$ELSE}
VH^.AnaCroisaxe := EstComptaPackAvance and GetParamSocSecur('SO_CROISAXE',False);
result          := VH^.AnaCroisaxe ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 06/06/2005
Modifié le ... : 18/08/2005
Description .. : Indique si il existe un paramétrage dans la base pour
Suite ........ : l'utilisation des tables libres de TIERSCOMPL
Suite ........ : - CA - 18/08/2005 : changement ZLI par ZLT suite modif
Suite ........ : Gescom pour Multisoc
Mots clefs ... :
*****************************************************************}
function TL_TIERSCOMPL_Actif : boolean;
begin
  Result := ExisteSQL ('SELECT CC_LIBELLE,CC_ABREGE,CC_CODE From CHOIXCOD '+
            ' Where CC_TYPE="ZLT" and cc_libelle <> ".-" And '+
            ' (CC_CODE Like "CT%"  or cc_code like "CR%") Order by CC_CODE');
end;

// ajout me Déplacement de CPGeneraux_TOM
procedure MajLettrageEcriture(Mp,Cpte : String ; Lefb : TFichierBase) ;
var
  LeWhere : String ;
begin
Case Lefb of                                               // ajout me 01/09/2005
     fbGene :LeWhere:='WHERE E_GENERAL="'+Cpte+'"' + ' AND E_ETATLETTRAGE ="RI"' ;  // CA - 22/04/2003 - On fait la manip sur toutes les écritures sinon pb dans rapport erreur ( compte lettrable avec état letrage incohérent ).
     fbAux  :LeWhere:='WHERE E_AUXILIAIRE="'+Cpte+'"' + 'AND E_ETATLETTRAGE ="RI"' ; // CA - 22/04/2003 - idem préc.
    End ;

{JP 03/08/05 : FQ 15990 : Gestion de E_TRESOSYNCHRO}
ExecuteSql('UPDATE ECRITURE SET E_MODEPAIE = "' + Mp + '", E_DATEECHEANCE = E_DATECOMPTABLE, ' +
           'E_ETATLETTRAGE = "AL", E_NUMECHE = 1, E_ECHE = "X", E_PAQUETREVISION = 1, E_TRESOSYNCHRO = "' +
           'CRE' + '" ' + LeWhere)
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/09/2005
Modifié le ... :   /  /
Description .. : Retourne TRUE si la chaîne passée en paramètre contient
Suite ........ : un caractère Joker SQL
Mots clefs ... :
*****************************************************************}
function EstJoker(St: string): boolean;
begin
  Result := False;
  if (pos('%', St) > 0)
    or (pos('[', St) > 0)
      or (pos('_', St) > 0)
        or (pos(']', St) > 0)
          or (pos('?', St) > 0)
            or (pos('*', St) > 0) then Result := True;
end;

{$IFDEF COMPTA}
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPEstTvaActivee : Boolean;
begin
  Result := GetParamSocSecur('SO_CPPCLSAISIETVA', False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/01/2006
Modifié le ... : 04/01/2007
Description .. : - LG - 04/01/2007 - FB 19930 - nouvelle fct 'tenue en
Suite ........ : recettes-depenses'
Mots clefs ... :
*****************************************************************}
function CPGetMontantTVA(vStCodeTva : string ; vDateDebut, vDateFin : TDateTime) : Double;
var lQuery    : TQuery;
    lTauxTva  : Double;
    lStWhere  : string;
    lStAxeTVa : string;
    lBoTvaSurEnc : Boolean;
begin
  Result := 0;
  if not GetParamSocSecur('SO_CPPCLSAISIETVA', False)  then Exit;

  lStAxeTva  := GetParamSocSecur('SO_CPPCLAXETVA', '');
  if lStAxeTva = '' then
  begin
    lStAxeTva := GetColonneSQL('AXE', 'X_AXE', 'X_LIBELLE = "TVA"');
    SetParamSoc('SO_CPPCLAXETVA', lStAxeTva);
  end;

  // Contrôle OK, on calcule

  lBoTvaSurEnc := GetParamSocSecur('SO_OUITVAENC', False);

  // Recherche du Taux dans CHOIXDPSTD
  lTauxTva := Valeur(GetColonneSql('CHOIXDPSTD', 'YDS_LIBRE', 'YDS_TYPE = "TVA" AND ' +
                            'YDS_CODE = "' + vStCodeTva + '" AND ' +
                            'YDS_NODOSSIER = "' + NoDossierBaseCommune + '" AND ' +
                            'YDS_PREDEFINI = "CEG"'));

  lStWhere := 'Y_SECTION = "' +  vStCodeTva + '" AND ' +
              'Y_AXE = "' + lStAxeTva + '" AND ' +
              'Y_DATECOMPTABLE >= "' + UsDateTime(vDateDebut) + '" AND ' +
              'Y_DATECOMPTABLE <= "' + UsDateTime(VDateFin) + '"' ;

  try
    lQuery := OpenSQL('SELECT SUM(Y_CREDIT)-SUM(Y_DEBIT) TOTAL FROM ANALYTIQ WHERE ' + lStWhere, True,-1,'',true);

    Result := lQuery.FindField('TOTAL').AsFloat;

    if lBoTvaSurEnc and (not GetParamSocSecur('SO_CPTVARECDEP',False)) then
      Result := Result / (1 + (lTauxTva /100));

  finally
    Ferme(lQuery);
  end;
end;

// AJOUT ME Pour les synchronisations
Procedure MAJHistoparam (NomTable, Data :string);
var
TB : TOB;
begin
   if (_IsDossierNetExpert or (GetParamsocSecur ('SO_CPMODESYNCHRO', false) = True)) then
   begin
         TB := TOB.Create ('CPHISTOPARAM', nil,-1);
         TB.Putvalue ( 'CHT_NOMTABLE' , NomTable);
         TB.Putvalue ( 'CHT_CODE' ,  Data);
         TB.InsertOrUpdateDB(TRUE);
         TB.free;
   end;
end;
//Spécificités agricoles
function CPEstComptaAgricole : Boolean;
begin
// MVG 11/10/2006 FQ 18948 RSI => RS
// MVG 28/03/2007 FQ 1987
 Result := ExisteSql ('SELECT DFI_REGIMFISCDIR FROM DPFISCAL Where DFI_REGLEFISC="BA" and (DFI_REGIMFISCDIR="RN" or DFI_REGIMFISCDIR="RS") and '+
 'DFI_GUIDPER="'+GetGuidPer(V_PGI.NoDossier)+'"' );
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/07/2007
Modifié le ... :   /  /    
Description .. : Test la présence d'ETAFI (ETAFI installé sur le poste)
Mots clefs ... : 
*****************************************************************}
function CPPresenceEtafi : boolean;
var stFileName : string;
begin
{$IFDEF EAGLCLIENT}
  stFileName := 'eEtafiS5.exe';
{$ELSE}
  stFileName := 'EtafiS5.exe';
{$ENDIF}
  Result := FileExists(ExtractFilePath(Application.ExeName) + '\' + stFileName);
end;

{$IFDEF COMSX}
// AJOUT ME Pour les synchronisations
Procedure MAJHistoparam (NomTable, Data :string);
{$IFNDEF EAGLSERVER}
var
TB : TOB;
{$ENDIF}
begin
{$IFNDEF EAGLSERVER}
   if (_IsDossierNetExpert or (GetParamsocSecur ('SO_CPMODESYNCHRO', false) = True)) then
   begin
         TB := TOB.Create ('CPHISTOPARAM', nil,-1);
         TB.Putvalue ( 'CHT_NOMTABLE' , NomTable);
         TB.Putvalue ( 'CHT_CODE' ,  Data);
         TB.InsertOrUpdateDB(TRUE);
         TB.free;
   end;
{$ENDIF}
end;
{$ENDIF}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/09/2006
Modifié le ... :   /  /
Description .. : Fonction de traçage des événements
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure CPEnregistreLOG( vStCodeLog : string ); // GCO - 18/09/2006
begin
  if ExecuteSQL('INSERT INTO LOG (LG_UTILISATEUR, LG_MENU, LG_DATE, LG_TEMPS, LG_COMMENTAIRE) ' +
    'VALUES ("'+ V_Pgi.User + '",' + IntToStr(FMenuG.GetLastNumCLick) + ',' + '"' +
     UsTime(Now) + '", "' + UsDateTime(iDate1900) + '", "' + VStCodeLog + '")') <> 1 then
  begin
    PgiInfo('Erreur SQL en insertion du traçage des évènements', '');
  end;
end;          


{$IFDEF CERTIFNF}
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : Nouvelle fonction de traçage des évènements
Mots clefs ... : 
*****************************************************************}
function CPEnregistreJalEvent( TypeEvent : String ; Libelle : String ; BlocNote : String ) :  integer;
var
  Q      : TQuery;
  NumEvt : Integer;
  SQL    : String;
begin
   Result := 0;
   if Trim(BlocNote) <> '' then
   begin
      // On récupere l'id de l'evenement :
      NumEvt := 0;
      Q := OpenSQL ('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1,'',true);
      try
         if not Q.EOF then
            NumEvt := Q.Fields[0].AsInteger;
      finally
         Ferme (Q);
      end;
      Inc(NumEvt);
      Result := NumEvt;
      SQL := 'INSERT INTO JNALEVENT ' +
             '(GEV_NUMEVENT, GEV_TYPEEVENT, GEV_LIBELLE, GEV_DATEEVENT, GEV_UTILISATEUR, GEV_ETATEVENT, GEV_BLOCNOTE) ' +
             'VALUES ( ' + IntToStr(NumEvt) + ', "' + TypeEvent + '", "' + Libelle + '", "' + UsDateTime(now) + '", "' + V_PGI.User + '", "OK", "' + BlocNote + '")';
      try
         ExecuteSQL(SQL);
      except
         // Il y a eu un soucis
         on E: Exception do
         begin
            PgiError('Erreur durant CPEnregistreJalEvent : ' + E.Message );
         end;
      end;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : Nouvelle fonction de traçage des évènements
Mots clefs ... : 
*****************************************************************}
function CPEnregistreJalEvent( TypeEvent : String ; Libelle : String ; BlocNote : HTStringList ) : integer ;
var
  Q      : TQuery;
  NumEvt : Integer;
  SQL    : String;
begin    
   Result := 0;
   if BlocNote.Count > 0 then
   begin
      // On récupere l'id de l'evenement :
      NumEvt := 0;
      Q := OpenSQL ('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1,'',true);
      try
         if not Q.EOF then
            NumEvt := Q.Fields[0].AsInteger;
      finally
         Ferme (Q);
      end;
      Inc(NumEvt);   
      Result := NumEvt;

      SQL := 'INSERT INTO JNALEVENT ' +
             '(GEV_NUMEVENT, GEV_TYPEEVENT, GEV_LIBELLE, GEV_DATEEVENT, GEV_UTILISATEUR, GEV_ETATEVENT, GEV_BLOCNOTE) ' +
             'VALUES ( ' + IntToStr(NumEvt) + ', "' + TypeEvent + '", "' + Libelle + '", "' + UsDateTime(now) + '", "' + V_PGI.User + '", "OK", "' + BlocNote.Text + '")';
      try
         ExecuteSQL(SQL);
      except
         // Il y a eu un soucis
         on E: Exception do
         begin
            PgiError('Erreur durant CPEnregistreJalEvent : ' + E.Message );
         end;
      end;
   end;
end;
{$ENDIF}

{$ENDIF}

////////////////////////////////////////////////////////////////////////////////

{b FP 21/02/2006}
{ TCompensation }

class function TCompensation.GetChampPlan: String;
begin
  Result := '';
  if IsCompensation then
    Result := 'T_CORRESP'+IntToStr(GetNumPlanCorrespondance);
end;

class function TCompensation.GetJalCompensation: String;
begin
  Result := GetParamSocSecur('SO_CPJALCOMPENSATION', '');
end;

class function TCompensation.GetNomPlanCorrespondance(NatureCpt, Compte: String): String;
var
  NumPlan: Integer;
  SQL:     String;
  Q:       TQuery;
begin
  Result := '';
  if not IsCompensation then
    Exit;
  if IsPlanCorresp(1) then
    NumPlan := 1
  else
    NumPlan := 2;

  SQL := 'SELECT CR_CORRESP'+
         ' FROM CORRESP'+
         ' WHERE CR_TYPE="AU'+IntToStr(NumPlan)+'"'+
         '   AND '+GetSQLCorrespCompte(NatureCpt, Compte);
  Q := OpenSQL(SQL, True,-1,'',true);
  if not Q.EOF then
    Result := Q.FindField('CR_CORRESP').AsString;
  Ferme(Q);
end;

class function TCompensation.GetNumPlanCorrespondance: Integer;
begin
  Result := 0;
  if IsPlanCorresp(1) then
    Result := 1
  else if IsPlanCorresp(2) then
    Result := 2;
end;

class function TCompensation.GetSQLCorrespCompte(NatureCpt, Compte: String): String;
begin
  if IsNatureClient(NatureCpt) then
    Result := ' CR_LIBRETEXTE1="'+Compte+'"'
  else if IsNatureFournisseur(NatureCpt) then {Fournisseur}
    Result := ' CR_LIBRETEXTE2="'+Compte+'"'
  else
    Result := ' 1<>1';
end;

class function TCompensation.IsCompensation: Boolean;
begin
  Result := IsPlanCorresp(1) or IsPlanCorresp(2);
end;

class function TCompensation.IsNatureClient(Nature: String): Boolean;
begin
  Result := (Nature='CLI') or (Nature='AUD');
end;

class function TCompensation.IsNatureFournisseur(Nature: String): Boolean;
begin
  Result := (Nature='FOU') or (Nature='AUC');
end;

class function TCompensation.IsPlanCorresp(NumPlan: Integer): Boolean;
begin
  Result := GetParamSocSecur('SO_CPGESTCOMPENSATION', False);
  if Result then
    Result := IntToStr(NumPlan) = GetParamSocSecur('SO_CPPLANCOMPENSATION', '');
end;
{e FP 21/02/2006}
function TZInfoCpta.local: boolean;
begin
  result := (FDossier = '') or ( FDossier = V_PGI.SchemaName ) ;
end;

{$IFNDEF NOVH}
{$IFDEF EAGLSERVER}
{ TLaVariableHalley }
constructor TLaVariableHalley.Create;
begin
  New(FVH);
end;
(*
destructor TLaVariableHalley.Destroy;
begin
  ReleaseLaVariableHalley();
  Dispose(FVH);
  inherited;
end;
*)
destructor TLaVariableHalley.Destroy;
begin
ReleaseLaVariableHalley (FVH);
Dispose(FVH);
FVH := nil;
inherited;
end;


{ LaVariableHalley en eAGLServer & Plugin métier :
  > les variables globales sont reroutées vers une fonction
  > Lors du premier appel à "VH", cette fonction "registre" la variable dans la session utilisateur serveur puis l'initialise
  > Lors d'autres utilisations, cette variable est récupérer telle quelle dans la session utilisateur }
function VH: PLaVariableHalley;
var
  Index: Integer;
  Session: TISession;
  IVH: TLaVariableHalley;
begin
  Result := nil;
  Session := LookupCurrentSession();
  if Assigned(Session) then
  begin
    Index := Session.UserObjects.IndexOf('VH');
    if Index >= 0 then
    begin
      IVH := TLaVariableHalley(Session.UserObjects.Objects[Index]);
      if Assigned(IVH) then
        Result := IVH.FVH
    end
    else
    begin
      IVH := TLaVariableHalley.Create();
      Session.UserObjects.AddObject('VH', IVH); //registrer avant l'appel de l'init sinon on boucle
      Result := IVH.FVH;
      InitLaVariableHalley();
      ChargeMagHalley();
    end
  end;
end;
{$ENDIF EAGLSERVER}
{$ENDIF !NOVH}

Function _GetParamSocSecur(St : String  ; ValDefaut : Boolean ; AvecMess : Boolean = FALSE) : Boolean ;
BEGIN
  {$IFDEF ENTREPRISE}
  Result:=FALSE ;
  {$IFNDEF NOVH}
  {$IFNDEF EAGLSERVER}
  if GetParamSocSecur('SO_CPANODYNA',false) then
    BEGIN
    Result:=TRUE ;
    If VH^.EnCours.EtatCpta<>'CPR' Then
      BEGIN
      Result:=FALSE ;
      If AvecMess Then PgiInfo('Vous devez effectuer une clôture provisoire pour que cette option soit active','A-Nouveaux dynamiques') ;
      END ;
    END Else Result:=FALSE ;
  {$ENDIF}
  {$ENDIF}
  {$ELSE}
  Result:=GetParamSocSecur('SO_CPANODYNA',false) ;
  {$ENDIF}

END ;

// GP le 20/08/2008 21496
Procedure RetoucheHVCPoursaisie(TC : tControl) ;
Var St : hString ;
    GoModif,OkEtab,OkJal : Boolean ;
    HVC : thvalcombobox ;
BEGIN
If TC=NIL Then Exit ;
If TC Is THValComboBox Then HVC:=thvalcombobox(TC) Else Exit ;
St:=AnsiUpperCase(HVC.DataType) ;
OkJal:=FALSE ; If GetParamSocSecur('SO_JALLOOKUP','') Then OkJal:=TRUE ;
OkEtab:=FALSE ; If GetParamSocSecur('SO_ETABLOOKUP','') Then OkEtab:=TRUE ;
GoModif:=FALSE ;
If (St='TTJALSAISIE') And OkJal Then GoModif:=TRUE ;
If (St='TTETABLISSEMENT') And OkEtab Then GoModif:=TRUE ;
If (St='CPJOURNAL') And OkJAL Then GoModif:=TRUE ;
If (St='TTJALBANQUE') And OkJAL Then GoModif:=TRUE ;
If (St='TTJALSANSECART') And OkJAL Then GoModif:=TRUE ;
If (St='TTJALFOLIO') And OkJAL Then GoModif:=TRUE ;
If (St='TTJOURNAUX') And OkJAL Then GoModif:=TRUE ;
If (St='TTJALEFFET') And OkJAL Then GoModif:=TRUE ;
If (St='CPJALSAISIEPARAM') And OkJAL Then GoModif:=TRUE ;
If (St='TTJALNONANOUVEAU') And OkJAL Then GoModif:=TRUE ;
If GoModif then HVC.DataType:=HVC.DataType+'2' ;
END ;

function OkMultiEntite: Boolean;
begin
  Result:=FALSE ;
  If Not GetParamSocSecur('SO_ME_ACTIF',FALSE) Then  Exit ;
  Result := True;
end;

function Entite : integer;
(* 0 si mono entité
   >0 si multi entité et sélectionné
   -2 si multi entité mais pas sélectionné.
*)
begin
  Result := 0;
  if OkMultiEntite then begin
  {$IFDEF CBP900}
    Result := TEntityModel.GetCurrentEntity;
    if Result=0 then
      Result:=-2;
  {$ENDIF}
    end;
end;

function GestionSequence:boolean;
begin
  Result:=V_PGI.NumVersionBase > 946;
end;


function CPIncrementerCompteur(TypeSouche,CodeSouche:string;DD:tDateTime;var MasqueNum:string17;Nombre:integer=1):integer;
var Req : string;
    wReq : string;
    lQuery : tQuery;
    Num : integer;
begin
  if Entite<0 then Result:=-2
  else if Nombre<0 then Result:=-3
  else begin
    Req := 'SELECT SH_NUMDEPART, SH_NUMDEPARTS, SH_SOUCHEEXO, SH_MASQUENUM FROM SOUCHE';
    wReq:= ' WHERE SH_SOUCHE="'+CodeSouche+'" AND '+
           ' SH_TYPE="'+TypeSouche+'" AND '+
           ' SH_ENTITY='+IntToStr(Entite);
    try
      Result:=-4;
      lQuery:=OpenSql(Req+wReq,true);
      if lQuery.EOF then Result:=-1
      else begin
        MasqueNum:=lQuery.FindField('SH_MASQUENUM').AsString;
        if GestionSequence then
        begin
          Req:=CleSequence(TypeSouche,CodeSouche,DD,lQuery.FindField('SH_SOUCHEEXO').AsString='X');
          try
            Result:=GetNextSequence(Req,Nombre,Entite);
          except
            raise Exception.Create(traduirememoire('Problème d''accès au compteur de séquence'));
          end;
        end;
      end;
    finally
      Ferme(lQuery);
    end;
  end;
end;

function CodeExo(DD:tDateTime):string;
var i : integer;
begin
  Result:=GetEnCours.Code ;
  If (dd>=GetEnCours.Deb) and (dd<=GetEnCours.Fin) then Result:=GetEnCours.Code
  else If (dd>=GetSuivant.Deb) and (dd<=GetSuivant.Fin) then Result:=GetSuivant.Code
  else If (dd>=GetPrecedent.Deb) and (dd<=GetPrecedent.Fin) then Result:=GetPrecedent.Code
  else For i:=1 To 5 do begin
    If (dd>=GetExoClo[i].Deb) And (dd<=GetExoClo[i].Fin)then Result:=GetExoClo[i].Code;
  end;
end;

procedure LibelleEtablissement(CodeEtab: string; var lib1, lib2: string);
var qq    : TQuery;
    StSQL : string;
begin

  lib1 := '';
  lib2 := '';

  STSQL := 'SELECT ET_LIBELLE,ET_ABREGE FROM ETABLISS WHERE ET_ETABLISSEMENT="' + CodeEtab + '"';

  QQ := OpenSQL(StSQL, false);

  if not QQ.EOF then
  begin
    lib1 := QQ.findField('ET_LIBELLE').asString;
    lib2 := QQ.findField('ET_ABREGE').asString;
  end;

  ferme(QQ);

end;

procedure LibelleDomaine(CodeDom: string; var lib1, lib2: string);
var qq    : TQuery;
    StSQL : string;
begin

  lib1 := '';
  lib2 := '';

  STSQL := 'SELECT BTD_LIBELLE FROM BTDOMAINEACT WHERE BTD_CODE="' + CodeDom + '"';

  QQ := OpenSQL(StSQL, false);

  if not QQ.EOF then
  begin
    lib1 := QQ.findField('BTD_LIBELLE').asString;
    lib2 := '';
  end;

  ferme(QQ);

end;


initialization
RegisterAglProc('ChargeMagHalley',False,0,AGLChargeMagHalley) ;
{$IFNDEF NOVH}
{$IFNDEF EAGLSERVER}
RegisterAglProc('PositionneDomaineUser',True,2,AGLPositionneDomaineUser) ;
{$ENDIF}
{$ENDIF}
if Not IsLibrary then ;


{$IFNDEF NOVH}
finalization
{$IFNDEF HVCL}
if Not IsLibrary then  ;
 BEGIN
{$IFNDEF EAGLSERVER}
  ReleaseLaVariableHalley();
{$ENDIF}
 END ;
{$ENDIF}
{$ENDIF}


end.









