unit EntGC;

interface

uses
  Variants,
  SysUtils,
  HCtrls,
  Windows,
  HEnt1,
  {$IFNDEF EAGLCLIENT}
	  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
	    DB,
  {$ENDIF !EAGLCLIENT}
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      {$IFNDEF PGIMAJVER}
        {$IFDEF STK}
          wMnu,
        {$ENDIF STK}
      {$ENDIF PGIMAJVER}
    {$ENDIF !ERADIO}
  {$ENDIF EAGLSERVER}
  {$IFDEF NOMADE}
    uToxClasses,
  {$ENDIF NOMADE}
  Registry,
  Menus,
  HCrc,
  HMsgBox,
  Messages,
  {$IFNDEF PLUGIN}
    EntPGI,
    Ent1,
  {$ENDIF !PLUGIN}
  {$IFDEF AFFAIRE}
	  UAFO_Ressource, // PL le 09/02/06 : gestion de la ressource associée au user courant
    UAFO_Affaire, // BDU - 02/02/07. Dans \gescom\liba
  {$ENDIF AFFAIRE}
  {$IFDEF GCGC}
    wParamFonction,
  {$ENDIF GCGC}
  uEntCommun,
  UTOB
  ;

Type
    T_RechArt = (traAucun,traErreur,traErrFour,traOk,traGrille,traOkGene,traSus,traContreM,traSt,traCancel) ;
    T_TraitAcompte = (TtaNone,TtaNormal,TtaProrata,TtaReliquat) ;
    TEtatGestionUnite = (meaNormal, meaAdvanced, meaAdvConvU);
    TGestUniteMode = set of TEtatGestionUnite;

	TTVSTA2014 = class (TObject)
		TOBpiece : TOB;
    TOBSSTrait : TOB;
    assigned : Boolean;
  end;


  tCleGQ = record { DISPO }
    Depot   : String;   { Dépôt }
    Article : string;   { Article }
  end;

{$IFDEF STK}
  tCleGQD = record { DISPODETAIL }
    Depot         : string;   { Dépôt }
    Article       : string;   { Article }
    StatutDispo   : string;   { Statut de disponibilité }
    StatutFlux    : String;   { Statut de flux }
    Emplacement   : string;   { Emplacement }
    LotInterne    : string;   { Lot interne }
    SerieInterne  : string;   { Série interne }
    TiersProp     : string;   { Tiers propriétaire }
    IndiceArticle : string;   { Indice article }
    Marque        : string;   { Marque }
    ChoixQualite  : string;   { Choix de qualité }
    RefAffectation: string;   { Référence affectation }
    Fournisseur   : string;   { Fournisseur contremarque }
    Reference     : string;   { Référence fournisseur contremarque }
  end;

  tCleGSM = record
    StkTypeMVT: string;
    QualifMVT : String;
    Guid      : string;
  end;

  tCleGEM = Record
    Depot       : string;
    Emplacement : string;
  end;

  tCleGSF = Record { Formule de lot/série }
    TypeGSF    : string;  { Type de formule }
    Depot      : string;  { Dépôt }
    FamilleNiv1: string;  { Famille article niveau 1 }
    FamilleNiv2: string;  { Famille article niveau 2 }
    FamilleNiv3: string;  { Famille article niveau 3 }
    Collection : string;  { Collection article }
    LibreArt1  : string;  { Table libre article 1 }
    LibreArt2  : string;  { Table libre article 2 }
    LibreArt3  : string;  { Table libre article 3 }
    LibreArt4  : string;  { Table libre article 4 }
    LibreArt5  : string;  { Table libre article 5 }
    LibreArt6  : string;  { Table libre article 6 }
    LibreArt7  : string;  { Table libre article 7 }
    LibreArt8  : string;  { Table libre article 8 }
    LibreArt9  : string;  { Table libre article 9 }
    LibreArtA  : string;  { Table libre article 10 }
    Article    : string;  { Article }
  end;

  tCleGST = Record  { Fiche de lot/série }
    TypeGST     : string;   { Type de fiche }
    Article     : string;   { Article }
    LotInterne  : string;   { Lot interne }
    SerieInterne: string;   { Série interne }
  end;

  tRefOrigine = record { référence d'origine (pièce, ligne, ordre, phase, besoin, ...) }
    PrefixeOri : string;  { Préfixe de la table d'origine }
    NatureOri  : string;  { Nature de l'origine }
    SoucheOri  : string;  { Souche d'origine }
    NumeroOri  : integer; { Numéro d'origine }
    IndiceOri  : integer; { Indice d'origine }
    OpeCircOri : string;  { Phase d'origine }
    NumLigneOri: integer; { Ligne d'origine }
  end;

  tCleGIA = Record          { Indices article }
    Article: string;        { Article }
    IndiceArticle: string;  { Indice article }
  end;
{$ENDIF STK}

  tOptionDuplic = record { options de duplication des pièces }
    { commandes marché }
    ChangeCodeMarche : boolean;
    CodeMarche : string;
    ChangeDate : boolean;
    DateDeb : TDateTime;
    DateFin : TDateTime;
    DupliqueDepot : boolean;
    CleDocOri : R_CleDoc;
  end;

{$IFDEF GCGC}
Type R_CLEAFFAIRE = RECORD
                NbPartie,Co1Lng,Co2Lng,Co3Lng : Integer ;
                Co1Type,Co2Type,Co3Type : String3 ;
                Co1valeur,Co2Valeur,Co3Valeur,Co1Lib,Co2Lib,Co3Lib,
                Co1valeurPro,Co2ValeurPro,Co3ValeurPro : String ;
                Co2Act,Co3Act,Co2Visible,Co3Visible,ProDifferent,GestionAvenant : Boolean;
                END;

Const
  SO_gcGestUniteMode = 'SO_GCGESTUNITEMODE';

  sNatureAuxiFournisseur  = 'FOU';
  sNatureAuxiClient       = 'CLI';
  sNatureAuxiProspect     = 'PRO'; 

const
  ASizeRADIO = 3;
  {$IFDEF GPAO}
    {$IFDEF EDI}     ASizeEDI     = 9; {$ELSE EDI}     ASizeEDI     = 0; {$ENDIF EDI}
    {$IFDEF CAISSE}  ASizeCAISSE  = 1; {$ELSE CAISSE}  ASizeCAISSE  = 0; {$ENDIF CAISSE}
    {$IFDEF AFFAIRE} ASizeAFF     = 2; {$ELSE AFFAIRE} ASizeAFF     = 0; {$ENDIF AFFAIRE}
    {$IFDEF GRC}
      {$IFNDEF GRCLIGHT}
        ASizeGRC = 1
      {$ELSE   GRCLIGHT}
        ASizeGRC = 0
      {$ENDIF !GRCLIGHT}
      {$IFDEF CTI}
// CRM_20080901_MNG_FQ;012;10918
      {$IFNDEF GRCLIGHT}
      + 1
      {$ENDIF GRCLIGHT}
      {$ENDIF CTI};
    {$ELSE GRC}
      ASizeGRC = 0;
    {$ENDIF GRC}
    { Constantes de la sérialisation GPAO }
    GPAO_GP_S5    = 1;  { Manufacturing }
    GPAO_CF_ST_S5 = 2;  { Configurateur standard }
    GPAO_CF_AV_S5 = 3;  { Configurateur avancé }
    GPAO_SAV_S5   = 4;  { SAV }
    GPAO_EGC_S5   = 5;  { eCommerce }
    GPAO_GRF_S5   = 6;  { Relation Fournisseur }
    GPAO_SCM_S5   = 7;  { SCM }
    GPAO_OBJ_S5   = 8;  { SCM }
    GPAO_RAD_S5   = 9;  { Radio }
    GPAO_EDI_S5   = GPAO_RAD_S5 + ASizeRADIO + 1; { EDI }
    GPAO_QUA_S5   = GPAO_EDI_S5 + ASizeEDI   + 1; { Qualité }

    {Taille et accès aux éléments du tableau des n° de sérialisation GPAO}
    ASizeGPAO_Seul = 11 + ASizeRADIO + ASizeEDI;
    ASizeGPAO = ASizeGPAO_Seul
              + ASizeCAISSE
              + ASizeAFF
              + ASizeGRC
              ;
              
    {$IFDEF CAISSE}
      PosCAISSE = ASizeGPAO_Seul + 1;
    {$ELSE  CAISSE}
      PosCAISSE = -1;
    {$ENDIF CAISSE}
    {$IFDEF AFFAIRE}
      PosGA = ASizeGPAO_Seul + ASizeCAISSE + 1;
    {$ELSE  AFFAIRE}
      PosGA = -2;
    {$ENDIF AFFAIRE}
    {$IFDEF GRC}
      {$IFNDEF GRCLIGHT}
        PosGRC = ASizeGPAO_Seul + ASizeCAISSE + ASizeAFF + 1;
        PosCTI = PosGRC + 1;
      {$ELSE   GRCLIGHT}
        PosGRC = -3;
        PosCTI = ASizeGPAO_Seul + ASizeCAISSE + ASizeAFF + 1;
      {$ENDIF !GRCLIGHT}
    {$ELSE  GRC}
      PosGRC = -4;
      PosCTI = -5;
    {$ENDIF GRC}
  {$ENDIF GPAO}
  RADIO_VTE = 1;  { Indice pour code séria "transaction création et transform° de pièce vente" }
  RADIO_ACH = 2;  { Indice pour code séria "transaction création et transform° de pièce achat" }
  RADIO_STK = 3;  { Indice pour code séria "transaction type stock" }

  { code client pour clients spécifiques}
  { Utilise le paramsoc SO_AFCLIENT }
  cInClient = 0;          // standard pas de numero de client
  cInClientAlgoe = 1;     // algoe
  cInClientOnyx = 2;      // onyx - otus - pas arrondi pour ville de paris
  cInClientAmyot = 3;     // Amyot
  cInClientDirEnergy = 4; // Direct Energy
  cInClientRenosol = 5;   // Renosol - revision de prix calculées à partir du dernier prix
  cInCLientRwd = 6;       // Rwd : tableau de bord adapté pour gestion d'un montant global p
  cInClientHygieneOffice = 7; // Hygiene Office a le droit de saisir des décimals en attendant la gestion en heure dans le générateur
  cInClientKPMG = 8; // pour KPMG
  cInClientBuffalo = 9; // Buffalo Grill
  cInClientSudEstNett = 11; // IsudEst Nettoyage (Etaient 5 à l'époque) GA_20080523_GME_GA15238

{ Paramètre d'appel de la procédure SaisiePiece }
Type R_SaisiePieceParam = record
      CodeTiers: string;
      Affaire: string;
      Avenant: string;
      PieceATraiter: string;
      SaisieAvanc: boolean;
	    OrigineEXCEL: Boolean;
      PasBouclerCreat: Boolean;
      NumOrdre: integer;
      NoPersp: Integer;
      RCdeOuv: Boolean;
  end;

type
  LaVariableGC = Class{$IFDEF EAGLSERVER}(TLaVariable){$ENDIF EAGLSERVER}
     Private
     Public
     // Générique
     TOBParPiece,TOBGCE,TOBGCB,TOBGCBDGD,TOBEdt,TOBRGDOC,TOBRGPRE,TOBPieceEDT : TOB ;
     MTobRestrFiche, MTobFicConf, MTobChpsOblFic, MTOBMEA, MTOBParPieceComp, MTOBParPieceDomaine : TMemoryTob;
     TobDevise: tMemoryTob;
     // Gescom
     ChargeMenu : boolean;
     GCOuvreTarifQte,GCArrondiLigne,GCValoArtDepot,
//      GCMontreNumero,GCDimCollection,GCVentCptaArt,GCVentCptaTiers,GCVentCptaAff,
     GCVentAxe2,GCVentAxe3,GCBloqueMarge,GCDefFactureHT,GCArticlesLies,GCMultiDepots,IsChpsOblig : boolean ;
     GCControleMarge : boolean; // DBR - Fiche 10603
     GCLgMinBarre,GCLgMinRefTiers,GCNbMoisStockOuverts : integer ;
     GCDepotDefaut,GCFamilleTaxe1,GCComportePrixDev,GCMargeArticle,GCModeRegleDefaut : String ;
     GCPhotoFiche,GCImpModeleDefaut,GCEcartConvert,GCAlerteDevise,GCPeriodeStock : String ;
     GCPCB, GCPRIXPOURQTE: Double ;
     GCDateClotureStock : TDateTime ;
     MTobChpsOblig, MGCTOBDim, MGCTOBAna : TMemoryTob;
     // Compta
     GCCpteEscVTE,GCCpteEscACH,GCCpteRemACH,GCCpteRemVTE,GCCpteHTVTE,GCCpteHTACH,
     GCCptePortACH,GCCptePortVTE,GCPontComptable,GCMarcheVentilAna : String ;
     GCLastRefPiece : String ;
     // Affaires
     CleAffaire : R_CleAffaire ;
     AFGereCritGroupeConf : string; //mcd 10/03/2005
     AfEquipe : string; //mcd 11/03/2005
     CleAppel : R_CleAffaire;
     AFGestionCom,AFRechAffAv,AFRechResAv,AFResCalculPR,AFProposAct,UserInvite : Boolean ;
     AFDateDebutAct,AFDateFinAct,AFDateAnalyseTB : TDateTime ;
     AFTypeSaisieAct,AFMesureActivite,AFNatAffaire,AFNatProposition,
{$ifdef GIGI}
     AfNatTiersGRCGI, AfNatTiersCOncGRCGI , AfTypTiersGRCGI,  // a initialiser ne fct nature gérer
{$endif}
       AFFReconduction,AFRepriseActiv,AFFGenerAuto,AFFLibfactAff,AFPrestationRes,AFFRacineAuxi,
       AFAcompte,AFTypevaloAct,AFFormatExer,AFProfilGener,AFValoActPR,AFValoActPV, RessourceUser : String ;
     AFValoActPrestPR,AFValoActPrestPV,AFValoActFraisPR,AFValoActFraisPV,AFValoActFourPR,AFValoActFourPV :string;
     AFTOBAffectChmp : TOB;
     AFTOBTraduction : TOB;
     AFTOBInterne : TOB; // mcd 26/02/03
{$IFDEF AFFAIRE}
     AFORessourceUser : TAFO_Ressources; // PL le 09/02/06 : pointeur sur la liste contenant la ressource associée au user courant : index = 0
     AFTobBlocageAff : TOB; // ST le 07/11/07 : Liste des blocages affaire, pour ne lire qu'une fois
{$ENDIF}
{$ifdef GIGI}
     AfTobCatalogue : TOB; //mcd 12/2005
{$endif}
     //seria
     BOSeria,GCSeria,ECSeria ,GRCSeria,GRFSeria,GAStdSeria,GAPlanningSeria,GAPlanChargeSeria,GAAchatSeria,GAFactInfo,GASegment, GASeria, GAIndusSeria, GASaisieDecSeria,VTCSeria : boolean ;
     PCPServeurSeria, SAVSeria, SCMSeria, OBJSeria  : Boolean;
                 
     // C.B 09/05/2006
     // seria des langues
     UKSeria, GERSeria, ITLSeria, ESPSeria, PORSeria, DUTSeria : boolean;
                          
     // BBI : GP Light
     OASeria : boolean;
     GCAchatStockSeria : boolean ;
     NbEtablisSeria  : integer ;
     // Back Office
     BOTypeMasque_Defaut : string ;
     // Front Office
     TOBPcaisse : TOB ;
     // Modif BTP
     BTGestParag,BTPrixMarche,BTSeriaAO, BTSeriaES, BTSeriaLogistique,BTSeriaContrat,BTSeriaIntervention,SeriaNEGOCE,
     BTSeriaGRC, BTSeriaChantiers,SeriaCoTraitance, SeriaSousTraitance,SeriaMobilite, BTSeriaParcMateriel,SeriaPlanCharge : boolean;
     BTTypeMoInterne : string;
     BSVUploadOK : Boolean;
     BSVOpenDoc : Boolean;
     ISGEDBSV : Boolean;
     BTCODESPECIF : string;
{$IFDEF BTP}
		 TOBParamTaxe : TOB;
     TOBArtEcart : TOB;
     TOBBPM : TOB;
     TOBTABFERIE : TOB;

     ModeValoPa : string;
     AutoLiquiTVAST : string;
{$ENDIF}
     {$IFDEF GPAO}
       GPAO_Seria: Array[1..ASizeGPAO] of Boolean;
     {$ENDIF GPAO}
     RADIOSeria: Boolean;
     RADIO_Seria: Array[1..ASizeRADIO] of Boolean;
     {$IFDEF QUALITE}
       QUALITESeria: Boolean;
     {$ENDIF QUALITE}
     GCIfDefCEGID : boolean ;
     ModeGestionEcartComptable : string; {DBR CPA}
     {$IFDEF NOMADE}
     PCPAchatSeria, PCPVenteSeria, PCPUsVte, PCPUsAch : boolean;
     PCPPrefixe, PCPRepresentant : string;
     PCPPceVente, PCPPceAchat : string;
     {$ENDIF}
     //CHR _20071220_GFO on sort la declaration du STK pour compiler wFormConv_Tom
     TobWWF : tMemoryTOB;  { Tob des formules de conversion }
//GP_20080616_MM_GP14941
     TobYFO : tMemoryTob;  { Tob des paramètres de tarifs }
     {$IFDEF STK }
       TobSDI : tMemoryTob;  { Tob des statuts de disponibilité }
       TobSFL : tMemoryTob;  { Tob des statuts de Flux }
       TobGC1 : tMemoryTob;  { Tob des compteurs de stock }
       TobGSN : tMemoryTob;  { Tob des nature de mouvement }
       TobGSR : tMemoryTob;  { Tob des règles de gestion }
       TobGVP : tMemoryTob;  { Tob des paramètre de valorisation }
       TobGVT : tMemoryTob;  { Tob des types de valorisation }
       TobTemp: Tob;  { Tob temporaire }
       TobWWA : tMemoryTOB;  { Tob des champs disponibles dans les formules de conversion }
       {$IFNDEF EAGLSERVER}
         {$IFNDEF ERADIO}
           {$IFNDEF PGIMAJVER}
             PmFlux : TPopupMenuFlux;  { Flux courant: STO; VTE; ACH; PRO; CON}
           {$ENDIF PGIMAJVER}
         {$ENDIF !ERADIO}
       {$ENDIF EAGLSERVER}
			 { Mis en place temporairement. Sert à revenir à l'ancien mode de gestion en cas de problème }
     {$ENDIF STK}
     GereTVAIntraComm : boolean; // JTR - TVA intracommunautaire
     TypeTVAIntraComm : string; // JTR - TVA intracommunautaire
     GestUniteMode: TEtatGestionUnite;
     NewLienPiece: Boolean;
     {$IFDEF CHR}
     ModeDeconnecte : Boolean;
     GCTOBEtab : TOB;
     LeSiteCourant : string; //LB - Journal des évènements
     {$ENDIF CHR}
     TOBTiersImpactPce, TOBElementImpactPce : TOB;
     TobTauxNegSoc : TOB;
     TobElementCptaPce : tMemoryTob; { Eléments d'une pièce impactant la compta en modification }
     RecupInfos : HTStrings;
     {$IFDEF AFFAIRE}
     VLibellesDesAffaires: TAFO_LibelleAffaire; // BDU - 02/02/07 - Objet conteneur des libellés des affaires
     VTiersParc: TAFO_TiersParc; // BDU - 04/2/04/07 - Objet conteneur des tiers du parc
     {$ENDIF}
      TobDupPceTiers : tMemoryTob;
      TobTypeCptaException : tMemoryTob;
     {$IFDEF GCGC}
       ParamFonction: TParamFonction;
//GP_20080625_DKZ_TD9702
       TobAJF       : tMemoryTOB;
     {$ENDIF GCGC}
     NumModuleCourant : Integer;
     TobNumGroupe : Tob;

     Procedure Initialisation;
     procedure BeforeProtecAFF(sender : TObject);
     Published
     END ;
{$IFDEF BTP}
	Type BTPTypeTaxe = (bttTVA,bttTPF,BttUnknown);
{$ENDIF}
Var
  TheParamTva : TTVSTA2014;
  {$IFNDEF EAGLSERVER}
    VH_GC : LaVariableGC ;
  {$ENDIF !EAGLSERVER}
  StWhereZoneLibre : string;
  {$IFDEF NOMADE}
    PCP_LesSites : TCollectionSites ; // Chargé au démarrage de l'appli.
  {$ENDIF NOMADE}

{$IFDEF GCGC}
{$IFDEF EAGLSERVER}
function VH_GC: LaVariableGC;
{$ENDIF EAGLSERVER}

{$IFDEF BTP}
procedure GetParamTaxe ;
procedure GetParamJF;
function  GetInfoParSouche ( SoucheG,Champ : String ) : Variant ;

Function SetParamTaxe : boolean;
function TypeTaxe (Indice : integer): BTPTypeTaxe; overload;
function BTTypeTaxe (CodeTaxe : string): BTPTypeTaxe;
{$ENDIF}
function  GetAfEquipe : string; //mcd 11/03/2005
{$ENDIF GCGC}

Function  GetInfoParPiece ( NaturePieceG,Champ : String ) : Variant ;
Function  GetInfoParPieceCompl ( NaturePieceG,Etablissement,Champ : String ) : Variant ;
Function  GetInfoParPieceDomaine ( NaturePieceG,Domaine,Champ : String ) : Variant ;
procedure ChargeArticlePiece; { GC_DBR_GC13811 }
Procedure ChargeDescriGC ;
Procedure ChargeParamsGC ;
Procedure InitLaVariableGC ;
{$IFDEF NOMADE}
Procedure InitLaVariableGCPCP;
{$ENDIF}
Procedure LibereLaVariableGC ;
Procedure AfInitAffectCat (Ismul : Boolean =true); 
procedure UpdateCombosGC ;
Function  GCGetTitreDim ( NumDim : integer ) : String ;
Function  GCGetIndiceDim ( Grille1,Grille2,Grille3,Grille4,Grille5 : String ; NumDim : integer ) : integer ;
Function  GCGetTitreDimRemplie ( Grille1,Grille2,Grille3,Grille4,Grille5 : String ; NumDim : integer ) : String ;
Function  GCGetCodeDim ( Grille,Code : String ; NumDim : integer ) : String ;
Function  GCGetCodeDimRemplie ( Grille1,Grille2,Grille3,Grille4,Grille5,
                                Code1,Code2,Code3,Code4,Code5 : String ; NumDim : integer ) : String ;
Function GCGetCodeDimORLIRemplie ( Grille1,Grille2,Grille3,Grille4,Grille5,
                                        Code1,Code2,Code3,Code4,Code5 : String ; NumDim : integer ) : String ;

Function OnChangeUserGC : boolean;
procedure GCTestDepotDefaut; // JTR - Dépôt par défaut
function AffecteDepotDefaut : string; // JTR - Dépôt par défaut
{$ENDIF}  //GCGC
Function  ForceEuroGC : boolean ;
Procedure TenteInverse ( TOBP : TOB ) ;
Procedure ForcePieceEuro(TOBPiece,TOBBases,TOBEches,TOBPorcs,TOBNomenclature,TOBAcomptes: TOB; TOBOuvrage:TOB=Nil;TOBPIeceRG:TOB=Nil;TOBBasesRG:TOB=Nil )  ;
function GetNaturePieceCde (ForSql : boolean=true) : string;
function GetNaturePieceLBT (ForSql : boolean=true) : string;
function GetNaturePieceBLC (ForSql : boolean=true) : string;
Function GetPieceAchat (WithRetour: boolean = True;WithCF : boolean=false;ForSql : boolean=true; WithRecepRegroupe : boolean=true;WithpropositionAchat : boolean=false): string;
procedure InitChampsSupSTORG;
{$IFDEF GCGC}
procedure ChargeParamsCegid ;
Procedure ChargeTobsConfid ;
{$IFDEF CHR}
procedure ChargeVHGCLeSiteCourant;  // LB - Journal des évènements
{$ENDIF CHR}
{$ENDIF}

{$IFDEF GPAO}
procedure GPCreationUNIDansMEA;
procedure GPVerificationQuotiteHHCC;
{$ENDIF GPAO}
         // fct utiliser pour les groupe de confidentialité
function GIGereCritGroupeConf : string;
function wGereConfigurator: Boolean;

{$IFDEF GCGC}
function GetParamSocGCGestUniteMode: String;
function GetParamSocGCDefQualifUniteGA: String;
function GetParamSocGCDefUniteGA: String;
function GetParamSocGCNewLienPiece: Boolean;
{$IFDEF CHR}
{$IFDEF EAGLCLIENT}
function GetInfoGCTOBEtab( Etablissement, NomChamp : string ) : string;
{$ENDIF EAGLCLIENT}
{$ENDIF CHR}
{$ENDIF GCGC}

function GereSTAGC: boolean;

function JAiLeDroitMenu122: Boolean;

{$IFDEF GIGI}
Procedure ChargeParamsocPerso;
{$ENDIF GIGI}

{$IFDEF STK}
function GetSqlGSN: string;
function GetSqlSDI: string;
function GetSqlGC1: string;
procedure InitVH_GCGestUnitMode;
{$ENDIF STK}

function GetSqlGDPTT : string;
function GetSqlGCXTT : string;
function GetCleDocByTob(TobXXX: Tob; const PrefixeSiPasTobReelle: String = ''): R_CleDoc;

Function ExisteSalarieModulePaie :Boolean;  // CCMX-CEGID ORGANIGRAMME DA
{$IFDEF GCGC} // BDU : Remplace GESCOM par GCGC
{$IFNDEF PAIEGRH}
{$IFNDEF PGIMAJVER}
{$IFNDEF CRM}
{$IFNDEF CMPGIS35}
  procedure InitOrganigrammePaie; // CCMX-CEGID ORGANIGRAMME DA
{$ENDIF}
{$ENDIF}
{$ENDIF !PGIMAJVER}
{$ENDIF}
{$ENDIF}
function GetSqlGCETT : string;
function GetInfoArtEcart (Article : String) : TOB;
function IsLivraisonClient (TOBL : TOB) : Boolean;
function IsLivraisonClientBTP (TOBL : TOB) : Boolean;
function IsLivraisonClientNEG (TOBL : TOB) : Boolean;
function GetNaturePieceNEG (ForSql : boolean=true) : string;
//
procedure InitDocumentTva2014;
procedure AssigneDocumentTva2014 (TOBpiece,TOBSSTrait : TOB);


implementation

Uses
  WinProcs,
  ShellAPI,
  HDimension,
  {$IFDEF GCGC}
     wCommuns,
    {$IFNDEF PGIMAJVER}
{$IFDEF CHR}
{$IFDEF EAGLCLIENT}
   TobUtil,
{$ENDIF EAGLCLIENT}
{$ENDIF CHR}
      {$IFNDEF CMPGIS35}
        UtilGC,
      {$ENDIF CMPGIS35}
    {$ENDIF PGIMAJVER}
    //mcd 04/12/2006 passe de gescom en GCGC
    {$IFNDEF PAIEGRH}
      {$IFNDEF OGC}
        {$IFNDEF PGIMAJVER}
          {$IFNDEF CRM}
{$IFNDEF CMPGIS35}
            EntPaie, // CCMX-CEGID ORGANIGRAMME PAIE POUR DEMANDES D'ACHATS SI PAS MODULE PAIE
{$ENDIF}
          {$ENDIF !CRM }
        {$ENDIF !PGIMAJVER}
      {$ENDIF OGC}
    {$ENDIF PAIEGRH}

    {$IF Defined(STK) and not Defined(PGIMAJVER)}
      MeaUtil,
    {$IFEND STK && !PGIMAJVER}
  {$ENDIF GCGC}
  {$IFNDEF PLUGIN}
  utilpgi,
  {$ENDIF PLUGIN}
  ParamSoc
   ,CbpMCD
   ,UconnectBSV
   ,CbpEnumerator
  ;


procedure InitDocumentTva2014;
begin
  TheParamTva.TOBpiece := nil;
  TheParamTVA.TOBSSTrait := nil;
  TheParamTVA.assigned := false;
end;

procedure AssigneDocumentTva2014 (TOBpiece,TOBSSTrait : TOB);
begin
  TheParamTva.TOBpiece := TOBPiece;
  TheParamTVA.TOBSSTrait := TOBSSTrait;
  TheParamTVA.assigned := True;
end;



{$IFDEF BTP}
procedure GetParamJF;
var Q : TQuery;
begin
  Q := OpenSQL('Select CO_LIBRE from COMMUN where CO_TYPE="TJF"', True);
  VH_GC.TOBTABFERIE.LoadDetailDB('COMMUN','','',Q,false);
  ferme (Q);  
end;


procedure GetParamTaxe ;
var QQ : TQuery;
begin
	VH_GC.TOBParamTaxe.ClearDetail;
  QQ := OpenSql ('SELECT * FROM PARAMTAXE',true,-1,'',true);
  VH_GC.TOBParamTaxe.LoadDetailDB ('PARAMTAXE','','',QQ,false,true);
  ferme (QQ);
end;

procedure SetInfosTaxes;
var Indice : integer;
    TOBTablette,UneTOB,TOBCC : TOB;
begin
  TOBTablette := TOB.Create ('LA TABLE',nil,-1);
  TRY
    if not VH_GC.TOBParamTaxe.UpdateDB (true) then BEGIN V_PGI.IOError := oeUnknown; Exit; End;
    for Indice := 0 to VH_GC.TOBParamTaxe.Detail.count -1 do
    begin
      UneTOB := VH_GC.TOBParamTaxe.Detail[Indice];
      TOBCC := TOB.Create ('COMMUN',TOBTablette,-1);
      TOBCC.PutValue('CC_TYPE','GCX');
      TOBCC.PutValue('CC_CODE',UneTOB.GetValue('BPT_CATEGORIETAXE'));
      TOBCC.PutValue('CC_LIBELLE',UneTOB.GetValue('BPT_LIBELLE'));
      TOBCC.PutValue('CC_ABREGE',Copy(UneTOB.GetValue('BPT_LIBELLE'),1,17));
    end;
    if not TOBTablette.UpdateDB (true) then BEGIN V_PGI.IOError := oeUnknown; Exit; End;
    AvertirTable ('GCCATEGORIETAXE')
  FINALLY
  TOBTablette.free;
  END;
end;

function SetParamTaxe : boolean;
begin
	result := false;
  BEGINTRANS;
  TRY
		SetInfosTaxes;
    COMMITTRANS;
  EXCEPT
  	ROLLBACK;
  END;
end;

function TypeTaxe (Indice : integer): BTPTypeTaxe;
begin
	result := BttUnknown;
	if Indice > VH_GC.TOBParamTaxe.Detail.count then exit;
	if VH_GC.TOBParamTaxe.detail[Indice-1].getValue('BPT_TYPETAXE') = 'TVA' then result := bttTVA else
	if VH_GC.TOBParamTaxe.detail[Indice-1].getValue('BPT_TYPETAXE') = 'TPF' then result := bttTPF;
end;

function BTTypeTaxe (CodeTaxe : string): BTPTypeTaxe;
var TheTOB : TOB;
begin
	result := BttUnknown;
  TheTOB := VH_GC.TOBParamTaxe.findfirst(['BPT_CATEGORIETAXE'],[CodeTaxe],true);
	if theTOB= nil then exit;
	if theTOB.getValue('BPT_TYPETAXE') = 'TVA' then result := bttTVA else
	if theTOB.getValue('BPT_TYPETAXE') = 'TPF' then result := bttTPF;
end;

{$ENDIF}


function GetSqlGDP : string;
begin
  Result := 'SELECT * FROM DOMAINEPIECE';
end;

function GetSqlGPC : string;
begin
  Result := 'SELECT * FROM PARPIECECOMPL';
end;

function GetSqlMEA : string;
begin
  Result := 'SELECT * FROM MEA';
end;

function GetSqlGOB_Champ : string;
begin
  Result := 'Select * from PARAMOBLIG where GOB_OBLIGATOIRE="X"';
end;

function GetSqlGOB_Fiche : string;
begin
  Result := 'Select * from PARAMOBLIG where GOB_GRVISIBLE<>"" or GOB_GRENABLE<>"" or GOB_USVISIBLE1<>"" or GOB_USENABLE1<>""';
end;

function GetSqlCO_POB (LaRequete : string) : string;
begin
  Result := LaRequete;
end;

function GetSqlANA (bAnalytique : boolean) : string;
begin
  if bAnalytique then Result := 'SELECT * FROM DECOUPEANA'
  else Result := 'SELECT * FROM STRUCRANAAFFAIRE';
end;

function GetSqlGDI : string;
begin
  Result := 'SELECT * FROM DIMENSION';
end;

Procedure TenteInverse ( TOBP : TOB ) ;
Var ic : integer ;
    Nam,NamC : String ;
BEGIN
for ic:=1 to TOBP.NbChamps do
    BEGIN
    NamC:=TOBP.GetNomChamp(ic) ;
    if Copy(NamC,Length(NamC)-2,3)='DEV' then
      BEGIN
      Nam:=Copy(NamC,1,Length(NamC)-3) ;
      if Copy(NamC,1,4)='GPE_' then Nam:=Nam+'ECHE' else
       if Copy(NamC,1,4)='GPB_' then Nam:=Nam+'TAXE' ;
      if TOBP.FieldExists(Nam) then TOBP.PutValue(NamC,TOBP.GetValue(Nam)) ;
      END ;
    END ;
END ;

Procedure ForcePieceEuro(TOBPiece,TOBBases,TOBEches,TOBPorcs,TOBNomenclature,TOBAcomptes: TOB; TOBOuvrage:TOB=Nil;TOBPIeceRG:TOB=Nil;TOBBasesRG:TOB=Nil)  ;
Var i,j : integer ;
BEGIN
TenteInverse(TOBPiece) ; TOBPiece.PutValue('GP_SAISIECONTRE','-') ;
for i:=0 to TOBPiece.Detail.Count-1 do BEGIN TenteInverse(TOBPiece.Detail[i]) ; TOBPiece.Detail[i].PutValue('GL_SAISIECONTRE','-') ; END ;
for i:=0 to TOBBases.Detail.Count-1 do BEGIN TenteInverse(TOBBases.Detail[i]) ; TOBBases.Detail[i].PutValue('GPB_SAISIECONTRE','-') ; END ;
for i:=0 to TOBEches.Detail.Count-1 do BEGIN TenteInverse(TOBEches.Detail[i]) ; TOBEches.Detail[i].PutValue('GPE_SAISIECONTRE','-') ; END ;
for i:=0 to TOBPorcs.Detail.Count-1 do TenteInverse(TOBPorcs.Detail[i]) ;
for i:=0 to TOBPorcs.Detail.Count-1 do for j:=0 to TOBPorcs.Detail[i].Detail.Count-1 do TenteInverse(TOBPorcs.Detail[i].Detail[j]) ;
for i:=0 to TOBAcomptes.Detail.Count-1 do TenteInverse(TOBAcomptes.Detail[i]) ;
// Modif BTP
if TOBOuvrage<>nil then for i:=0 to TOBOuvrage.Detail.Count-1 do TenteInverse(TOBOuvrage.Detail[i]) ;
if TOBPieceRG<>nil then for i:=0 to TOBPieceRG.Detail.Count-1 do TenteInverse(TOBPieceRG.Detail[i]) ;
if TOBBasesRG<>nil then for i:=0 to TOBBasesRG.Detail.Count-1 do TenteInverse(TOBBasesRG.Detail[i]) ;
// --
END ;

Function ForceEuroGC : boolean ;
BEGIN
{$IFDEF NOVH}
Result:=((GetParamSoc('SO_GCTOUTEURO')) or (V_PGI.DateEntree>=EncodeDate(2002,01,01))) ;
{$ELSE}
Result:=(VH^.TenueEuro) and ((GetParamSoc('SO_GCTOUTEURO')) or (V_PGI.DateEntree>=EncodeDate(2002,01,01))) ;
{$ENDIF NOVH}
END ;

{$IFDEF GCGC}
function GetInfoParPiece ( NaturePieceG,Champ : String ) : Variant ;
var
  TobNat : tob ;
begin
  result:='' ;
  TobNat := VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'],[NaturePieceG],False) ;
  if assigned(TobNat) then
  begin
    if Champ = 'GPP_LIBELLE' then
      result := TraduireMemoire(TobNat.GetValue(Champ))
    else
      result := TobNat.GetValue(Champ);
  end;
end;

Function GetInfoParPieceCompl ( NaturePieceG,Etablissement,Champ : String ) : Variant ;
Var TOBNat : TOB ;
BEGIN
Result:='' ;
VH_GC.MTobParPieceComp.Load;
TOBNat:=VH_GC.MTOBParPieceComp.FindFirst(['GPC_NATUREPIECEG','GPC_ETABLISSEMENT'],[NaturePieceG,Etablissement],False) ;
if TOBNat<>Nil then Result:=TOBNat.GetValue(Champ) ;
END ;

Function GetInfoParPieceDomaine ( NaturePieceG,Domaine,Champ : String ) : Variant ;
Var TOBNat : TOB ;
BEGIN
Result:='' ;

VH_GC.MTobParPieceDomaine.Load;

TOBNat:=VH_GC.MTOBParPieceDomaine.FindFirst(['GDP_NATUREPIECEG','GDP_DOMAINE'],[NaturePieceG,Domaine],False) ;
if TOBNat<>Nil then Result:=TOBNat.GetValue(Champ) ;
END ;
(*
Function GetRessourceUser : string;
var
  Q : TQuery;
begin
  result := '';
  Q := OpenSQL('Select ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE="'+V_PGI.User+'"',true,-1,'RESSOURCE');
  try
    if Not Q.EOF then result:=Q.fields[0].asString;
  finally
    ferme(Q);
  end;
end;
*)
Function TraiteUserinvite : boolean;
var
conf, champ : string;
i : integer;
begin
result:=False;
//Si superviseur , controleur ou utilisateur inconnu , on sort.
If (V_PGI.Superviseur) or (V_PGI.Controleur) or (VH_GC.RessourceUser='') then exit;
champ := '';
//On récupére le type de confidentialité
conf := GetParamsoc('SO_AFTYPECONF');
if conf = 'A01' then champ:='RESPONSABLE'
else
 begin
  For i:=1 to 3 do
   begin
    if conf = 'AS'+IntToStr(i) then champ:='RESSOURCE'+IntToStr(i);
   end;
  end;
//Si pas confidentialité par ressource libre ou par responsable on sort.
if champ = '' then exit;
// True si invité sinon false
Result :=  not ExisteSQL('Select AFF_AFFAIRE FROM AFFAIRE WHERE AFF_'+champ+'="'+VH_GC.RessourceUser+'"');
end;

Procedure ChargeParamsGC ;

  {$IFDEF STK}
  function GetSqlSFL: string;
  begin
    Result := 'SELECT CO_CODE AS SFL_STATUTFLUX'
            + ',Trim("#ICO#"||STR(CO_LIBRE)) AS SFL_IMAGE'
            + ' FROM COMMUN'
            + ' WHERE CO_TYPE="SFL"'
            + ' ORDER BY CO_CODE'
  end;
  {$ENDIF STK}

  {$IFDEF STK}
  function GetSqlGSR: string;
  begin
    Result := 'SELECT GSR_REGLEGESTION,GSR_LIBELLE,GSR_ORDRESELECT'
            + ' FROM STKREGLEGESTION'
            + ' ORDER BY GSR_REGLEGESTION'
  end;
  {$ENDIF STK}

  {$IFDEF STK}
  function GetSqlGVP: string;
  begin
    Result := 'SELECT *'
            + ' FROM STKVALOPARAM'
            + ' ORDER BY ISNULL(GVP_QUALIFMVT, " ") DESC,'
            +          ' ISNULL(GVP_STKFLUX, " ") DESC,'
            +          ' ISNULL(GVP_NATURETRAVAIL, " ") DESC,'
            +          ' ISNULL(GVP_FAMILLEVALO, " ") DESC,'
            +          ' ISNULL(GVP_FAMILLENIV1, " ") DESC,'
            +          ' ISNULL(GVP_FAMILLENIV2, " ") DESC,'
            +          ' ISNULL(GVP_FAMILLENIV3, " ") DESC'
  end;
  {$ENDIF STK}

  {$IFDEF STK}
  function GetSqlGVT: string;
  begin
    Result := 'SELECT *'
            + ' FROM STKVALOTYPE'
            + ' ORDER BY GVT_VALOTYPE'
  end;
  {$ENDIF STK}

  //CHR _20071220_GFO
  function GetSqlWWF: string;
  begin
    Result := 'SELECT * '
             + 'FROM WFORMCONV '
             + 'ORDER BY WWF_UNITEDEPART,WWF_UNITEARRIVEE,WWF_CODEFORMCONV';
  end;

  {$IFDEF STK}
  function GetSqlChoixCodWWA: string;
  begin
    Result := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="WWA" ORDER BY CC_CODE';
  end;
  {$ENDIF STK}

BEGIN
// PCH 08/02/2006 déplacement de l'appel à cette fonction
OnChangeUserGC;

// PCH 09/02/2006 personnalisation paramètres société
{$IFDEF GIGI}
  // voir le commentaire de fonctionnement à l'implémentation de cette procédure
  ChargeParamsocPerso;
{$ENDIF}

VH_GC.GCNbMoisStockOuverts:=GetParamSocSecur('SO_GCNBMOISSTOCKOUVERTS', 0) ;
VH_GC.GCAlerteDevise:=GetParamSocSecur('SO_GCALERTEDEVISE', '') ;
VH_GC.GCEcartConvert:=GetParamSocSecur('SO_GCECARTCONVERT', '') ;
VH_GC.GCVentAxe2:=GetParamSocSecur('SO_GCVENTAXE2', False) ;
if not VH_GC.GCVentAxe2 then
  VH_GC.GCVentAxe3:=False
  else
  VH_GC.GCVentAxe3:=GetParamSocSecur('SO_GCVENTAXE3', False) ;
if VH_GC.GcMarcheVentilAna = '' then VH_GC.GcMarcheVentilAna:='ZZZ';   // mcd 08/01/03 cette fct est appelé après maj paramsoc ==> perte info mise dans Mdisp
if not (ctxaffaire in V_PGI.PgiContexte) then VH_GC.ChargeMenu := false; //mcd 28/03/2007 on appelle cette fct sur saisie paramsco.. donc pb. fait en GIA dans Afterprotec
VH_GC.GCBloqueMarge:=GetParamSocSecur('SO_GCBLOQUEMARGE', False) ;
VH_GC.GCControleMarge := GetParamSocSecur ('SO_GCCONTROLEMARGE', False); // DBR - Fiche 10603
VH_GC.GCArticlesLies:=GetParamSocSecur('SO_GCARTICLESLIES', False) ;
VH_GC.GCMultiDepots:=GetParamSocSecur('SO_GCMULTIDEPOTS', 'X') ;
VH_GC.GCDepotDefaut := AffecteDepotDefaut; // JTR - Dépôt par défaut
VH_GC.GCLgMinBarre:=GetParamSocSecur('SO_GCLGMINBARRE', 0) ;
VH_GC.GCLgMinRefTiers:=GetParamSocSecur('SO_GCLGMINREFTIERS', 0) ;
VH_GC.GCImpModeleDefaut:=GetParamSocSecur('SO_GCIMPMODELEDEFAUT', '') ;
VH_GC.GCModeRegleDefaut:=GetParamSocSecur('SO_GCMODEREGLEDEFAUT', '') ;
VH_GC.GCFamilleTaxe1:=GetParamSocSecur('SO_GCFAMILLETAXE1', '') ;
VH_GC.GCOuvreTarifQte:=GetParamSocSecur('SO_GCOUVRETARIFQTE', '') ;
VH_GC.GCArrondiLigne:=GetParamSocSecur('SO_GCARRONDILIGNE', False) ;
VH^.DefCatTVA:=GetParamSocSecur('SO_GCDEFCATTVA', '') ;
VH^.DefCatTPF:=GetParamSocSecur('SO_GCDEFCATTPF', '') ;
VH_GC.GCComportePrixDev:=GetParamSocSecur('SO_GCCOMPORTEPRIXDEV', '') ;
VH_GC.GCValoArtDepot:=GetParamSocSecur('SO_GCVALOARTDEPOT', False) ;
VH_GC.GCMargeArticle:=GetParamSocSecur('SO_GCMARGEARTICLE', '') ;
VH_GC.GCPhotoFiche:=GetParamSocSecur('SO_GCPHOTOFICHE', '') ;
VH_GC.GCDefFactureHT:=GetParamSocSecur('SO_GCDEFFACTUREHT', False) ;
VH_GC.GCPCB:=GetParamSocSecur('SO_GCPCB', 0) ;
// Liaison comptable
VH_GC.GCCpteEscVTE:=GetParamSocSecur('SO_GCCPTEESCVTE', '') ;
VH_GC.GCCpteEscACH:=GetParamSocSecur('SO_GCCPTEESCACH', '') ;
VH_GC.GCCpteRemVTE:=GetParamSocSecur('SO_GCCPTEREMVTE', '') ;
VH_GC.GCCpteRemACH:=GetParamSocSecur('SO_GCCPTEREMACH', '') ;
VH_GC.GCCpteHTVTE:=GetParamSocSecur('SO_GCCPTEHTVTE', '') ;
VH_GC.GCCpteHTACH:=GetParamSocSecur('SO_GCCPTEHTACH', '') ;
VH_GC.GCCptePortVTE:=GetParamSocSecur('SO_GCCPTEPORTVTE', '') ;
VH_GC.GCCptePortACH:=GetParamSocSecur('SO_GCCPTEPORTACH', '') ;
VH_GC.GCPontComptable:=GetParamSocSecur('SO_GCPONTCOMPTABLE', '') ;
// Affaires
VH_GC.AFRechResAv:=GetParamSocSecur('SO_AFRECHRESAV', False) ;
{$IFDEF BTP}
   VH_GC.CleAppel.NbPartie  := GetParamSoc('SO_APPCODENBPARTIE') ;
   VH_GC.CleAppel.Co1Lng    :=GetParamSoc('SO_APPCO1LNG') ;
   VH_GC.CleAppel.Co2Lng    :=GetParamSoc('SO_APPCO2LNG') ;
   VH_GC.CleAppel.Co3Lng    :=GetParamSoc('SO_APPCO3LNG') ;
   VH_GC.CleAppel.Co1Type   :=GetParamSoc('SO_APPCO1TYPE') ;
   VH_GC.CleAppel.Co2Type   :=GetParamSoc('SO_APPCO2TYPE') ;
   VH_GC.CleAppel.Co3Type   :=GetParamSoc('SO_APPCO3TYPE') ;
   VH_GC.CleAppel.Co1valeur :=GetParamSoc('SO_APPCO1VALEUR') ;
   VH_GC.CleAppel.Co2Valeur :=GetParamSoc('SO_APPCO2VALEUR') ;
   VH_GC.CleAppel.Co3Valeur :=GetParamSoc('SO_APPCO3VALEUR') ;
   VH_GC.CleAppel.Co1valeurPro :=GetParamSoc('SO_APPCO1VALEURPRO') ;
   VH_GC.CleAppel.Co2ValeurPro :=GetParamSoc('SO_APPCO2VALEURPRO') ;
   VH_GC.CleAppel.Co3ValeurPro :=GetParamSoc('SO_APPCO3VALEURPRO') ;
   VH_GC.CleAppel.Co2Act    :=GetParamSoc('SO_APPCO2ACT') ;
   VH_GC.CleAppel.Co3Act    :=GetParamSoc('SO_APPCO3ACT') ;
   VH_GC.CleAppel.Co2Visible:=GetParamSoc('SO_APPCO2VISIBLE') ;
   VH_GC.CleAppel.Co3Visible:=GetParamSoc('SO_APPCO3VISIBLE') ;
   VH_GC.CleAppel.ProDifferent:=GetParamSoc('SO_AFFPRODIFFERENT') ;
   VH_GC.CleAppel.GestionAvenant:=GetParamSoc('SO_AFFGESTIONAVENANT') ;
   VH_GC.CleAppel.Co1Lib    :=GetParamSoc('SO_APPCO1LIB');
   VH_GC.CleAppel.Co2Lib    :=GetParamSoc('SO_APPCO2LIB');
   VH_GC.CleAppel.Co3Lib    :=GetParamSoc('SO_APPCO3LIB');
{$ENDIF}
if ((V_PGI.PGIContexte = [] ) or (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte)) then
   BEGIN //mcd 22/12/2005 passage en paramsocsecur
   VH_GC.CleAffaire.NbPartie  := GetParamSocSecur('SO_AFFCODENBPARTIE', 0) ;
   VH_GC.CleAffaire.Co1Lng    :=GetParamSocSecur('SO_AFFCO1LNG', 0) ;
   VH_GC.CleAffaire.Co2Lng    :=GetParamSocSecur('SO_AFFCO2LNG', 0) ;
   VH_GC.CleAffaire.Co3Lng    :=GetParamSocSecur('SO_AFFCO3LNG', 0) ;
   VH_GC.CleAffaire.Co1Type   :=GetParamSocSecur('SO_AFFCO1TYPE', '') ;
   VH_GC.CleAffaire.Co2Type   :=GetParamSocSecur('SO_AFFCO2TYPE', '') ;
   VH_GC.CleAffaire.Co3Type   :=GetParamSocSecur('SO_AFFCO3TYPE', '') ;
   VH_GC.CleAffaire.Co1valeur :=GetParamSocSecur('SO_AFFCO1VALEUR', '') ;
   VH_GC.CleAffaire.Co2Valeur :=GetParamSocSecur('SO_AFFCO2VALEUR', '') ;
   VH_GC.CleAffaire.Co3Valeur :=GetParamSocSecur('SO_AFFCO3VALEUR', '') ;
   VH_GC.CleAffaire.Co1valeurPro :=GetParamSocSecur('SO_AFFCO1VALEURPRO', '') ;
   VH_GC.CleAffaire.Co2ValeurPro :=GetParamSocSecur('SO_AFFCO2VALEURPRO', '') ;
   VH_GC.CleAffaire.Co3ValeurPro :=GetParamSocSecur('SO_AFFCO3VALEURPRO', '') ;
   VH_GC.CleAffaire.Co1Lib    :=GetParamSocSecur('SO_AFFCO1LIB', '');
   VH_GC.CleAffaire.Co2Lib    :=GetParamSocSecur('SO_AFFCO2LIB', '');
   VH_GC.CleAffaire.Co3Lib    :=GetParamSocSecur('SO_AFFCO3LIB', '');
   VH_GC.CleAffaire.Co2Act := GetParamSocSecur('SO_AFFCO2ACT', False);
   VH_GC.CleAffaire.Co3Act    :=GetParamSocSecur('SO_AFFCO3ACT', False) ;
   VH_GC.CleAffaire.Co2Visible:=GetParamSocSecur('SO_AFFCO2VISIBLE', False) ;
   VH_GC.CleAffaire.Co3Visible:=GetParamSocSecur('SO_AFFCO3VISIBLE', False) ;
   VH_GC.CleAffaire.ProDifferent:=GetParamSocSecur('SO_AFFPRODIFFERENT', False) ;
   VH_GC.CleAffaire.GestionAvenant:=GetParamSocSecur('SO_AFFGESTIONAVENANT', False) ;
   VH_GC.AFGestionCom         :=GetParamSocSecur('SO_AFGESTIONCOM', False) ;
   Vh_GC.afGereCritGroupeConf  := GiGereCritGroupeConf;
   VH_GC.AFRechAffAv          :=GetParamSocSecur('SO_AFRECHAFFAV', False) ;
   // VH_GC.AFRechResAv          :=GetParamSoc('SO_AFRECHRESAV') ;  mis quelque soit le contexte pour GC et GRC
   VH_GC.AFResCalculPR        :=GetParamSocSecur('SO_AFRESCALCULPR', False) ;
   VH_GC.AFProposAct          :=GetParamSocSecur('SO_AFPROPOSACT', False) ;
   VH_GC.AFDateFinAct         :=GetParamSocSecur('SO_AFDATEFINACT', iDate2099) ;
   VH_GC.AFDateAnalyseTB      :=GetParamSocSecur('SO_AFDATEANALYSETB', iDate1900) ;
   VH_GC.AFTypeSaisieAct      :=GetParamSocSecur('SO_AFTYPESAISIEACT', 'MOI') ;
   VH_GC.AFMesureActivite     :=GetParamSocSecur('SO_AFMESUREACTIVITE', '') ;
{$IFDEF BTP}  // MODIF BRL 27/01/05 : marre des paramètres pourris !!!
  SetParamsoc ('SO_AFNATAFFAIRE','DBT');
  SetParamsoc ('SO_AFNATPROPOSITION','ETU');
{$ENDIF}
   VH_GC.AFNatAffaire         :=GetParamSocSecur('SO_AFNATAFFAIRE', 'AFF') ;
   VH_GC.AFNatProposition     :=GetParamSocSecur('SO_AFNATPROPOSITION', 'PAF') ;
   VH_GC.AFFReconduction      :=GetParamSocSecur('SO_AFFRECONDUCTION', 'TAC') ;
   VH_GC.AFRepriseActiv       :=GetParamSocSecur('SO_AFREPRISEACTIV', 'MON') ;
   VH_GC.AFFGenerAuto         :=GetParamSocSecur('SO_AFFGENERAUTO', 'FOR') ;
   VH_GC.AFFLibfactAff        :=GetParamSocSecur('SO_AFFLIBFACTAFF', '') ;
   VH_GC.AFPrestationRes      :=GetParamSocSecur('SO_AFPRESTATIONRES', '') ;
   VH_GC.AFFRacineAuxi        :=GetParamSocSecur('SO_AFFRACINEAUXI', '') ;
   VH_GC.AFAcompte            :=GetParamSocSecur('SO_AFACOMPTE', '') ;
   VH_GC.AFTypevaloAct        :=GetParamSocSecur('SO_AFTYPEVALOACT', '001') ;
   VH_GC.AFFormatExer         :=GetParamSocSecur('SO_AFFORMATEXER', 'AUC') ;
//
//
   VH_GC.AFValoActPrestPR     :=GetParamSocSecur('SO_AFVALOACTPR', 'RES') ;  //Prix de revient des prestations
   VH_GC.AFValoActPrestPV     :=GetParamSocSecur('SO_AFVALOACTPV', 'ART') ;  //Prix de vente des prestations
   VH_GC.AFValoActFraisPR     :=GetParamSocSecur('SO_AFVALOFRAISPR', 'RES') ;//Prix de revient des frais
   VH_GC.AFValoActFraisPV     :=GetParamSocSecur('SO_AFVALOFRAISPV', 'ART') ;//Prix de vente des frais
   VH_GC.AFValoActFourPR      :=GetParamSocSecur('SO_AFVALOFOURPR', 'RES') ; //Prix de revient des fournitures
   VH_GC.AFValoActFourPV      :=GetParamSocSecur('SO_AFVALOFOURPV', 'ART') ; //Prix de vente des fournitures
   END ;

VH_GC.GCIfDefCEGID:=GetParamSocSecur('SO_IFDEFCEGID', False) ;
// PCH 08/02/2006
//OnChangeUserGC;

VH_GC.AfEquipe             := GetAfEquipe ; //mcd 11/03/2005 il faut que l'utilisateur soit chargé
ChargeTobsConfid;
//CHR _20071220_GFO
VH_GC.TobWWF := tMemoryTob.Create('WWF', GetSqlWWF);
{$IFDEF STK }
  { Tob des natures de mouvement }
  VH_GC.TobSDI := tMemoryTob.Create('SDI', GetSqlSDI);
  VH_GC.TobSFL := tMemoryTob.Create('SFL', GetSqlSFL);
  VH_GC.TobGC1 := tMemoryTob.Create('GC1', GetSqlGC1);
  VH_GC.TobGSN := tMemoryTob.Create('GSN', GetSqlGSN);
  VH_GC.TobGSR := tMemoryTob.Create('GSR', GetSqlGSR);
  VH_GC.TobGVP := tMemoryTob.Create('GVP', GetSqlGVP);
  VH_GC.TobGVT := tMemoryTob.Create('GVT', GetSqlGVT);
  VH_GC.TobWWA := tMemoryTob.Create('WWA', GetSqlChoixCodWWA);
  { Gestion des unités : Formule de conversion }
  InitVH_GCGestUnitMode;
{$ENDIF STK}

  // JTR - TVA intracommunautaire
{$IFNDEF CHR}
  VH_GC.GereTVAIntraComm := (GetParamSocSecur('SO_GERETVAINTRACOMM', False));
  VH_GC.TypeTVAIntraComm := (GetParamSocSecur('SO_TYPETVAINTRACOMM', ''));
{$ENDIF CHR}
  // Fin JTR
{$IFDEF GCGC}
  VH_GC.NewLienPiece := GetParamSocGCNewLienPiece;
{$ENDIF GCGC}
{$IFDEF BTP}
  VH_GC.BTTypeMoInterne:=GetParamSocSecur('SO_BTMOINTERNE','SAL;INT');
  if GetParamsUserBSV (VH_GC.BSVUploadOK,VH_GC.BSVOpenDoc) then VH_GC.ISGEDBSV := True;
  VH_GC.BTCODESPECIF := GetParamSocSecur('SO_BTCODESPECIF','');
// CHargement Des parametre des Taxes
GetParamTaxe;
GetParamJF;
{$ENDIF}
{$IFDEF CHR}
  ChargeVHGCLeSiteCourant; // LB - Journal des évènements
{$ENDIF CHR}
  VH_GC.TOBTiersImpactPce := TOB.Create('IMPACT TIERS SUR PIECE', nil, -1);
  TOB.Create('', VH_GC.TOBTiersImpactPce, -1);
  TOB.Create('', VH_GC.TOBTiersImpactPce, -1);
  VH_GC.TOBElementImpactPce := TOB.Create('', nil, -1);
  VH_GC.RecupInfos := HTStringList.Create;
  {$IFDEF GCGC} // BDU : Remplace GESCOM par GCGC
  {$IFNDEF PAIEGRH}
  {$IFNDEF PGIMAJVER}
  {$IFNDEF CRM}
{$IFNDEF CMPGIS35}
  InitOrganigrammePaie; // CCMX-CEGID ORGANIGRAMME DA
{$ENDIF}
  {$ENDIF}
  {$ENDIF !PGIMAJVER}
  {$ENDIF}
  {$ENDIF}
  VH_GC.TobTauxNegSoc := TOB.Create('TAUX NEGOCIE SOCIETE', nil, -1);
  VH_GC.TobDupPceTiers := tMemoryTob.Create('GDP', GetSqlGDPTT);
  VH_GC.TobElementCptaPce := tMemoryTob.Create('GCE', GetSqlGCETT);
  VH_GC.TobTypeCptaException := tMemoryTob.Create('GCE', GetSqlGCXTT);
  VH_GC.ModeValoPa := GetParamSocSecur('SO_BTVALOACHNSTOCK','PAA');
  VH_GC.AutoLiquiTVAST := GetParamSocSecur('SO_CODETVALIQUIDST','');
END ;

procedure ChargeParamsCegid ;
begin
  if VH_GC.GCIfDefCEGID then
  begin
    V_PGI.QRPDFOptions:=5 ;
    V_PGI.LookUpLocate:=True ;
  //$$$JP 25/03/04: n'existe plus  V_PGI.ListeByUser:=True ;
  end;
end;

procedure GCTestDepotDefaut; // JTR - Dépôt par défaut
var Qry : TQuery;
    Depot, NomTablette, Valeur : string;
    NumCpt : integer;
    StListe : HTStringList;
begin
  if (GetParamSoc('SO_GCDEPOTDEFAUT') = '') and
     (not ((CtxScot in V_PGI.PGIContexte) or (CtxTempo in V_PGI.PGIContexte) or (CtxChr in V_PGI.PGIContexte))) then
  begin
    Qry := OpenSql('SELECT GDE_DEPOT FROM DEPOTS',true,-1, '', True);
    if not Qry.Eof then
    begin
      NomTablette := 'GCDEPOT';
      NumCpt := TTTONum(NomTablette);
      if V_PGI.DECombos[NumCpt].Valeurs = nil then RemplirListe(NomTablette,'');
      StListe := V_PGI.DECombos[NumCpt].Valeurs;
      // Affecte dépôt qui correspond à l'établissement par défaut si existe
      // Sinon, affecte le 1er trouvé
      if GetParamSoc('SO_ETABLISDEFAUT') <> '' then
      begin
        for NumCpt := 0 to StListe.count -1 do
        begin
          Valeur := copy(StListe.Strings[NumCpt],1,pos(#9,StListe.Strings[NumCpt])-1);
          if Valeur = GetParamSoc('SO_ETABLISDEFAUT') then
          begin
            Depot := Valeur;
            break;
          end;
        end;
        if Depot = '' then
//GP_20071126_TP_GP14593
          Depot := copy(StListe.Strings[0],1,pos(#9,StListe.Strings[0])-1);
      end
      else
        Depot := copy(StListe.Strings[0],1,pos(#9,StListe.Strings[0])-1);
      SetParamSoc('SO_GCDEPOTDEFAUT',Depot);
      // Maj journal évènements et message
      {$IFNDEF PGIMAJVER}
      {$IFNDEF CMPGIS35}
      MAJJnalEvent ('', 'OK', 'Affectation du dépôt par défaut.', 'Le dépôt '+Depot+' a été affecté comme dépôt par défaut.');
      PgiInfo(Format(TraduireMemoire('Le dépôt : %s a été affecté comme dépôt par défaut dans les paramètres sociétés.'), [Depot]), 'Affectation du dépôt par défaut');
      {$ENDIF CMPGIS35}
      {$ENDIF PGIMAJVER}
      if VH_GC.GCDepotDefaut = '' then
        VH_GC.GCDepotDefaut := AffecteDepotDefaut;
    end else
      Ferme(Qry);
  end;
end;

function AffecteDepotDefaut : string; // JTR - Dépôt par défaut
var
  DepotUtilisat : string ;
begin
  Result := VH^.ProfilUserC[prEtablissement].Depot;
  DepotUtilisat := VH^.ProfilUserC[prEtablissement].Depot ;
  if Result = '' then // Si encore vide, affecte ParamSoc
  begin
    Result := GetParamSocSecur('SO_GCDEPOTDEFAUT', '') ;
  end;
end;

Procedure AfInitAffectCat(IsMul :Boolean=true);
{$ifdef GIGI}
Var
 stSql: string ;
{$endif}
begin
{$ifdef GIGI}
        //mcd 12/2005 chargement tob pour Catalogue service
  if VH_GC.AFTOBCatalogue<>Nil then VH_GC.AFTOBCatalogue.Free ;
  VH_GC.AFTOBCatalogue:=TOB.Create('',Nil,-1) ;
  If GetParamSocSecur('SO_AFCATALOGUE',False) then
    begin
      //on doit tout prendre... peu de choses
			//si
		if IsMul then stsql:= 'SELECT * FROM AFINFOGRCDP WHERE ADP_TYPEINFO="CAT" ORDER BY ADP_ONGLET ASC,ADP_PANEL DESC'
     else stsql:= 'SELECT * FROM AFINFOGRCDP WHERE ADP_TYPEINFO="CAT" ORDER BY ADP_ONGLET,ADP_PANEL';
		VH_GC.AFTOBCatalogue.LoadDetailFromSQL (stSql);
    end;
{$endif}
end;

Procedure AfInitAffectChmp; //mcd 21/11/2005  IE
 //permet d'alimenter la tob pour la gestion des affectation automatique ressource/ mul,qr1,cube ..
 //plus tob pour catalogue service
Var stSql: string ;
{$IFDEF AFFAIRE}
 ii : integer;
 TobDet : tob;
{$ENDIF AFFAIRE}
 (* QQ : TQuery;  *)
begin

  AfInitAffectCat;
     //seulement 2 champs, on peut tout charger
  Stsql := 'SELECT * FROM YLIENTABLETTE WHERE YLT_ORIGINETAB like "AFTLIBRERES%"';
  if VH_GC.AFTOBAffectChmp<>Nil then VH_GC.AFTOBAffectChmp.Free ;
{$IFNDEF BTP}
{$ifdef AFFAIRE}
  if (VH_GC.AFORessourceUser = nil) then
    begin
    FreeAndNil(Vh_GC.AftobAffectChmp);		exit // Ajout PL le 13/02/06
		end;
  VH_GC.AFTOBAffectChmp:=TOB.Create('',Nil,-1) ;
  Vh_GC.AftobAffectChmp.LoadDetailFromSQL (stSql);
  If Vh_GC.aftobAffectChmp.detail.count=0 then
    begin
    FreeAndNil(Vh_GC.AftobAffectChmp);
    exit;
    end;
  (* mcd 13/02/06 on ne refait pas lecture
     QQ:=Opensql ('SELECT ARS_LIBRERES1, ARS_LIBRERES2,ARS_LIBRERES3,ARS_LIBRERES4,ARS_LIBRERES5,ARS_LIBRERES6,ARS_LIBRERES7,ARS_LIBRERES8,ARS_LIBRERES9,ARS_LIBRERESA FROM RESSOURCE WHERE ARS_RESSOURCE="'
      + VH_GC.RessourceUser +'"',true);
  if Not QQ.EOF then
    begin  *)
  for ii:=0 to   Vh_GC.aftobAffectChmp.detail.count -1 do
    begin
    TobDet := Vh_GC.aftobAffectChmp.detail[ii];
    StSql := COpy (TobDet.Getvalue('YLT_ORIGINETAB'),12,1); //prend N° à la fin de AFTLibreRes* ...
      //mcd 13/02/02 TobDet.addChampSupValeur ('VALEUR', QQ.FindField('ARS_LIBRERES'+StSql).AsString);
    TobDet.addChampSupValeur ('VALEUR',VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetString('ARS_LIBRERES'+StSql));
    end;
{$endif}
{$ENDIF}
   (* end
  else FreeAndNil(Vh_GC.AftobAffectChmp); // pas de ressource associé, on efface la tob
  Ferme(QQ);  *)
end;

procedure GCInitChpsOblig; // JTR test champs obligatoire pièce
{$IFNDEF PGIMAJVER}
var
  TobTmp, TobTmp1 : TOB;
  Cpt, VenteOuAchat, PieceOuLigne, NumTable, Numchamp : integer;
   Sql : string;
	Mcd : IMCDServiceCOM;
	Field     : IFieldCOM ;

{$ENDIF !PGIMAJVER}
begin

{$IFNDEF PGIMAJVER}
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();

{$IFNDEF CMPGIS35}
  VH_GC.MTobChpsOblig.ClearDetail;
  TobTmp := TOB.Create('PIECE VENTE',VH_GC.MTobChpsOblig,-1);
  TOB.Create('CHPS PIECE',TobTmp,-1);
  TOB.Create('CHPS LIGNE',TobTmp,-1);
  TobTmp := TOB.Create('PIECE ACHAT',VH_GC.MTobChpsOblig,-1);
  TOB.Create('CHPS PIECE',TobTmp,-1);
  TOB.Create('CHPS LIGNE',TobTmp,-1);
{$IFDEF AFFAIRE}
    // création des info pour natures associé affaire et propo
  TobTmp := TOB.Create('PIECE AFFAIRE',VH_GC.MTobChpsOblig,-1);
  TOB.Create('CHPS PIECE',TobTmp,-1);
  TOB.Create('CHPS LIGNE',TobTmp,-1);
  TobTmp := TOB.Create('PIECE PROPOSITION',VH_GC.MTobChpsOblig,-1);
  TOB.Create('CHPS PIECE',TobTmp,-1);
  TOB.Create('CHPS LIGNE',TobTmp,-1);
{$ENDIF}
  TobTmp := TOB.Create('CHPS OBLIG',nil,-1);
{$IFDEF AFFAIRE} //mcd 08/03/04 ..
  Sql :='SELECT * FROM COMMUN LEFT JOIN PARAMOBLIG '+
                 'ON GOB_CODE=CO_CODE AND GOB_OBLIGATOIRE="X" '+
                 'WHERE CO_TYPE="POB" ';
    // si affaire seule, ou dan sla GC, ce n'est pas les même codes ) utiliser
  if (ctxaffaire in V_PGI.PGIContexte) then // on doit avoir les codes pour pieces
    Sql :=Sql +'AND (CO_CODE in ("A12","A13","A14","A15","B08","B10","GP1","GP3")) '
  else
    if (ctxgcaff in V_PGI.PGIContexte) then
      Sql :=Sql +'AND (CO_CODE like "GP%" OR (CO_CODE > "A10" AND CO_CODE <="A15"))';
  Sql := Sql +' ORDER BY CO_ABREGE, CO_CODE';
{$ELSE}//mcd 08/03/04 ..
  Sql := 'SELECT * FROM COMMUN LEFT JOIN PARAMOBLIG '+
                 'ON GOB_CODE=CO_CODE AND GOB_OBLIGATOIRE="X" '+
                 'WHERE CO_TYPE="POB" AND (CO_ABREGE="VEN" OR CO_ABREGE="ACH") '+
                 'ORDER BY CO_ABREGE, CO_CODE';
{$ENDIF} //mcd 08/03/04 ..
  TobTmp.LoadDetailDBFromSql ('', Sql);
  for Cpt := 0 to TobTmp.Detail.count -1 do
  begin
    if TobTmp.detail[Cpt].GetValue('CO_ABREGE') = 'VEN' then
      VenteOuAchat := 0
      else if TobTmp.detail[Cpt].GetValue('CO_ABREGE') = 'ACH' then
      VenteOuAchat := 1
{$IFDEF AFFAIRE}
      else if TobTmp.detail[Cpt].GetValue('CO_ABREGE') = 'AFF' then
      VenteOuAchat := 2  // cas pièce associé à l'affaire
      else if TobTmp.detail[Cpt].GetValue('CO_ABREGE') = 'PRO' then
      VenteOuAchat := 3  // cas pièce associée aux propostions
{$ENDIF}
      else
      VenteOuAchat := -1;
    if VenteOuAchat >= 0 then
    begin
{$IFDEF EAGLCLIENT}
      if (copy(TobTmp.detail[Cpt].GetString ('CO_LIBRE'),1,5) = 'PIECE') and (TobTmp.detail[Cpt].GetString ('GOB_CODE') <> '') and
         (TobTmp.detail[Cpt].GetString ('GOB_CODE') <> #0) then
        PieceOuLigne := 0
        else if (copy(TobTmp.detail[Cpt].GetString('CO_LIBRE'),1,5) = 'LIGNE') and (TobTmp.detail[Cpt].GetString('GOB_CODE') <> '') and
                (TobTmp.detail[Cpt].GetString ('GOB_CODE') <> #0) then
        PieceOuLigne := 1
        else
        PieceOuLigne := -1;
{$ELSE}
      if (copy(TobTmp.detail[Cpt].GetValue('CO_LIBRE'),1,5) = 'PIECE') and (TobTmp.detail[Cpt].GetString('GOB_CODE')<>'') then
                 //PCS 28122005 support AGL580  and (not VarIsNull(TobTmp.detail[Cpt].GetValue('GOB_CODE')))
        PieceOuLigne := 0
        else if (copy(TobTmp.detail[Cpt].GetValue('CO_LIBRE'),1,5) = 'LIGNE') and (TobTmp.detail[Cpt].GetString('GOB_CODE')<>'') then
                //PCS 28122005 support AGL580 and (not VarIsNull(TobTmp.detail[Cpt].GetValue('GOB_CODE')))
             PieceOuLigne := 1
             else
             PieceOuLigne := -1;
{$ENDIF EAGLCLIENT}
      if PieceOuLigne >= 0 then
      begin
        if VH_GC.IsChpsOblig = False then
          VH_GC.IsChpsOblig := True;
        TobTmp1 := TOB.Create('LE CHAMP',VH_GC.MTobChpsOblig.Detail[VenteOuAchat].Detail[PieceOuLigne],-1);
        TobTmp1.AddChampSupValeur('NOMCHAMP',TobTmp.detail[Cpt].GetValue('GOB_NOMCHAMP'),False);
//mcd 08/03/04 .. pour affaire et propo, idem VEN       if VenteOuAchat = 0 then
// DBR 27/12/04...donc il faut tester VenteOuAchat <> 1 car si <> 1 est autre que ACH donc VEN et tout le reste
        if VenteOuAchat <> 1 then
          TobTmp1.AddChampSupValeur('LIBCHAMP',LibTableLibreCliFour('VEN', TobTmp.detail[Cpt].GetValue('GOB_NOMCHAMP')),False)
          else
          TobTmp1.AddChampSupValeur('LIBCHAMP',LibTableLibreCliFour('ACH', TobTmp.detail[Cpt].GetValue('GOB_NOMCHAMP')),False);
        if PieceOuLigne = 0 then
          TobTmp1.AddChampSupValeur('NOMTABLE','PIECE',False)
          else
          TobTmp1.AddChampSupValeur('NOMTABLE','LIGNE',False);
        NumTable := TableToNum(TobTmp1.GetValue('NOMTABLE'));
        NumChamp := ChampToNum(TobTmp1.GetValue('NOMCHAMP'));
        Field := mcd.getField(TobTmp1.GetValue('NOMCHAMP')) ;

        if Field.tipe = 'DOUBLE' then
          TobTmp1.AddChampSupValeur('VALEURVIDE',0.0,False)
        else if Field.tipe = 'DATE' then
          TobTmp1.AddChampSupValeur('VALEURVIDE', iDate1900,False)
        else if Field.tipe = 'INTEGER' then
          TobTmp1.AddChampSupValeur('VALEURVIDE', 0,False)
        else if Field.tipe = 'BOOLEAN' then
          TobTmp1.AddChampSupValeur('VALEURVIDE', '-',False)
        else
          TobTmp1.AddChampSupValeur('VALEURVIDE', '',False);
      end;
    end;
  end;
  FreeAndNil(TobTmp);
{$ENDIF CMPGIS35}
{$ENDIF PGIMAJVER }
end;


Function GCGetTitreDim ( NumDim : integer ) : String ;
BEGIN
Result:=RechDom('GCCATEGORIEDIM','DI'+IntToStr(NumDim),True) ;    // Modif JCF pour prendre l'abrégé
END ;

Function GCGetIndiceDim ( Grille1,Grille2,Grille3,Grille4,Grille5 : String ; NumDim : integer ) : integer ;
Var GG : Array[1..5] of String ;
    i,Ind : integer ;
BEGIN
Result:=0 ; Ind:=0 ;
GG[1]:=Grille1 ; GG[2]:=Grille2 ; GG[3]:=Grille3 ; GG[4]:=Grille4 ; GG[5]:=Grille5 ;
for i:=1 to MaxDimension do if GG[i]<>'' then BEGIN Inc(Ind) ; if Ind=NumDim then BEGIN Result:=i ; Break ; END ; END ;
END ;

Function GCGetTitreDimRemplie ( Grille1,Grille2,Grille3,Grille4,Grille5 : String ; NumDim : integer ) : String ;
Var Ind : integer ;
BEGIN
Result:='' ;
Ind:=GCGetIndiceDim(Grille1,Grille2,Grille3,Grille4,Grille5,NumDim) ;
if Ind>0 then Result:=GCGetTitreDim(Ind) ;
END ;

Function GCGetCodeDim ( Grille,Code : String ; NumDim : integer ) : String ;
Var TOBD : TOB ;
BEGIN
Result:='' ;

VH_GC.MGCTobDim.Load;

TOBD:=VH_GC.MGCTOBDim.FindFirst(['GDI_TYPEDIM','GDI_GRILLEDIM','GDI_CODEDIM'],
                               ['DI'+IntToStr(NumDim),Grille,Code],False) ;
if TOBD<>Nil then Result:=TOBD.GetValue('GDI_LIBELLE') ;
END ;

Function GCGetCodeDimRemplie ( Grille1,Grille2,Grille3,Grille4,Grille5,
                               Code1,Code2,Code3,Code4,Code5 : String ; NumDim : integer ) : String ;
Var TOBD : TOB ;
    Ind  : integer ;
    GG,CC : Array[1..5] of String ;
BEGIN
Result:='' ;
Ind:=GCGetIndiceDim(Grille1,Grille2,Grille3,Grille4,Grille5,NumDim) ; if Ind<=0 then Exit ;
GG[1]:=Grille1 ; GG[2]:=Grille2 ; GG[3]:=Grille3 ; GG[4]:=Grille4 ; GG[5]:=Grille5 ;
CC[1]:=Code1   ; CC[2]:=Code2   ; CC[3]:=Code3   ; CC[4]:=Code4   ; CC[5]:=Code5   ;

VH_GC.MGCTobDim.Load;

TOBD:=VH_GC.MGCTOBDim.FindFirst(['GDI_TYPEDIM','GDI_GRILLEDIM','GDI_CODEDIM'],
                               ['DI'+IntToStr(Ind),GG[Ind],CC[Ind]],False) ;
if TOBD<>Nil then Result:=TOBD.GetValue('GDI_LIBELLE') ;
END ;

Function GCGetCodeDimORLIRemplie ( Grille1,Grille2,Grille3,Grille4,Grille5,
                               Code1,Code2,Code3,Code4,Code5 : String ; NumDim : integer ) : String ;
Var TOBD : TOB ;
    Ind  : integer ;
    GG,CC : Array[1..5] of String ;
BEGIN
Result:='' ;
Ind:=GCGetIndiceDim(Grille1,Grille2,Grille3,Grille4,Grille5,NumDim) ; if Ind<=0 then Exit ;
GG[1]:=Grille1 ; GG[2]:=Grille2 ; GG[3]:=Grille3 ; GG[4]:=Grille4 ; GG[5]:=Grille5 ;
CC[1]:=Code1   ; CC[2]:=Code2   ; CC[3]:=Code3   ; CC[4]:=Code4   ; CC[5]:=Code5   ;

VH_GC.MGCTobDim.Load;

TOBD:=VH_GC.MGCTOBDim.FindFirst(['GDI_TYPEDIM','GDI_GRILLEDIM','GDI_CODEDIM'],
                               ['DI'+IntToStr(Ind),GG[Ind],CC[Ind]],False) ;
if TOBD<>Nil then Result:=TOBD.GetValue('GDI_DIMORLI') ;
END ;

{ GC_DBR_GC13811_DEBUT }
procedure ChargeArticlePiece;
var
  TobNat, TobP, TobG : Tob;
  iInd : integer;
  Nat : string;
begin
  for iInd := 0 to VH_GC.TobParPiece.Detail.Count - 1 do
    VH_GC.TobParPiece.Detail [iInd].Detail.Clear;

  TobP := Tob.Create ('', nil, -1);
  try
    TOBP.LoadDetailFromSql ('SELECT * FROM ARTICLEPIECE');
    for iInd := TOBP.Detail.Count-1 downto 0 do
    begin
      TOBG := TOBP.Detail[iInd] ;
      Nat := TOBG.GetValue('GAP_NATUREPIECEG') ;
      TOBNat := VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'], [Nat], False) ;
      if TOBNat <> Nil then
        TOBG.ChangeParent (TOBNat,-1)
      else TobG.Free;
    end;
  finally
    TobP.Free;
  end;
end;
{ GC_DBR_GC13811_FIN }

procedure ReinitlesMemoryTObs;

  procedure ReinitMemTOb (TheMemTOB : tMemoryTob);
  begin
    TheMemTOB.ClearDetail;
    TheMemTOB.IsLoaded := false;
  end;
  
begin
  ReinitMemTOb(VH_GC.TobDevise);
  ReinitMemTOb(VH_GC.MTobParPieceDomaine);
  ReinitMemTOb(VH_GC.MTOBParPieceComp);
  ReinitMemTOb(VH_GC.MGCTobDim);
  ReinitMemTOb(VH_GC.MGCTobAna);
  ReinitMemTOb(VH_GC.MTobChpsOblig);
  ReinitMemTOb(VH_GC.MTobMEA);
end; 

procedure ChargeLesmemoryTOBs;
begin
  ReinitlesMemoryTObs;
  VH_GC.TobDevise.Load;
  VH_GC.MTobParPieceDomaine.load;
  VH_GC.MTOBParPieceComp.Load;
  VH_GC.MGCTobDim.load;
  VH_GC.MGCTobAna.load;
  VH_GC.MGCTOBAna.Detail.Sort('GDA_AXE;GDA_ETABLISSEMENT;GDA_TYPESTRUCRANA;GDA_TYPECOMPTE;GDA_RANG;');
  VH_GC.MTobChpsOblig.load;
  VH_GC.MTobMEA.Load;
end;

Procedure ChargeDescriGC ;
BEGIN
if VH_GC.MTOBMEA<>Nil then VH_GC.MTOBMEA.Free ;
if VH_GC.TOBParPiece<>Nil then VH_GC.TOBParPiece.Free ;
if VH_GC.MTOBParPieceComp<>Nil then VH_GC.MTOBParPieceComp.Free ;
if VH_GC.MTOBParPieceDomaine<>Nil then VH_GC.MTOBParPieceDomaine.Free ;
VH_GC.MTOBMEA := TMemoryTob.Create ('MEA', GetSqlMEA);
VH_GC.MTobParPieceDomaine := tMemoryTob.Create ('GDP', GetSqlGDP);
VH_GC.MTOBParPieceComp:=TMemoryTob.Create ('GPC', GetSqlGPC);
VH_GC.TOBParPiece:=TOB.Create('',Nil,-1);
VH_GC.TOBParPiece.LoadDetailFromSQL ('SELECT * FROM PARPIECE');
if VH_GC.TOBParPiece.Detail.Count<=0 then Exit ;
ChargeArticlePiece; { GC_DBR_GC13811 }
if not Assigned (VH_GC.MGCTobDim) then
  VH_GC.MGCTobDim := TMemoryTob.Create ('GDI', GetSqlGDI);
if not Assigned (VH_GC.MGCTobAna) then
  VH_GC.MGCTobAna := TMemoryTob.Create ('GDA', GetSqlANA (GetParamSoc ('SO_GCAXEANALYTIQUE')));
VH_GC.MTobChpsOblig := TMemoryTob.Create ('_CO_OBLIG_', '', GCInitChpsOblig);
AfInitAffectChmp; //mcd 21/11/2005  IE
{$IFDEF GPAO}
GPCreationUNIDansMEA;
GPVerificationQuotiteHHCC;
{$ENDIF GPAO}

{$IFDEF BTP}
ChargeLesmemoryTOBs;
{$ENDIF}
END ;

procedure UpdateCombosGC ;
Var ii,i : integer ;
    St : String ;
BEGIN
St:='TTTVA' ; ii:=TTToNum(St) ; if ii>0 then V_PGI.DECombos[ii].Where:='CC_TYPE="'+VH^.DefCatTVA+'"' ;
St:='TTTPF' ; ii:=TTToNum(St) ; if ii>0 then V_PGI.DECombos[ii].Where:='CC_TYPE="'+VH^.DefCatTPF+'"' ;
//Paul
for i:=1 to 5 do
    begin
    St:='GCGRILLEDIM'+intToStr(i) ;
    ii:=TTToNum(St) ;
    if ii>0 then V_PGI.DECombos[ii].Libelle:= RechDom('GCCATEGORIEDIM','DI'+IntToStr(i),FALSE) ;
    end;
if VH_GC.GCIfDefCEGID then
begin
  St:= 'GCREPRESENTANT';
  ii:=TTToNum(st) ;
  V_PGI.DECombos[ii].SaisieCode :=true;   //VH_GC.GCComLookUp;
  V_PGI.DECombos[ii].ChampLib:='GCL_COMMERCIAL||":"||GCL_LIBELLE';
end
else
begin
  if GetParamSoc('SO_GCCODEREPRESENTANT') then
     begin
     St:= 'GCREPRESENTANT';
     ii:=TTToNum(st) ;
     V_PGI.DECombos[ii].SaisieCode :=true;   //VH_GC.GCComLookUp;
     V_PGI.DECombos[ii].ChampLib:='GCL_COMMERCIAL||":"||GCL_LIBELLE';
     St:= 'GCCOMMERCIAL';
     ii:=TTToNum(st) ;
     V_PGI.DECombos[ii].SaisieCode :=true;   //VH_GC.GCComLookUp;
     V_PGI.DECombos[ii].ChampLib:='GCL_COMMERCIAL||":"||GCL_LIBELLE';
     V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' ORDER BY GCL_LIBELLE';
     end;
end;
{$IFDEF STK}
  St:= 'GCQUALIFMVT';
  ii:=TTToNum(St);
  V_PGI.DECombos[ii].Champ := '';

  St:= 'GCSTKQUALIFMVT';
  ii:=TTToNum(St);
  V_PGI.DECombos[ii].Champ := V_PGI.DECombos[ii].Champ + ';QUALIFMVT';
{$ENDIF STK}
END ;

procedure InitChampsSupSTORG;
begin
  VH_GC.TOBRGDOC.AddChampSupValeur ('NATUREPIECEG','');
  VH_GC.TOBRGDOC.AddChampSupValeur('SOUCHE','');
  VH_GC.TOBRGDOC.AddChampSupValeur('NUMERO',0);
  VH_GC.TOBRGDOC.AddChampSupValeur('INDICEG',0);
  VH_GC.TOBRGDOC.AddChampSupValeur('DATEMODIF',iDate1900);
end;

procedure InitChampsSupRGPRE;
begin
  VH_GC.TOBRGPRE.AddChampSupValeur ('NATUREPIECEG','');
  VH_GC.TOBRGPRE.AddChampSupValeur('SOUCHE','');
  VH_GC.TOBRGPRE.AddChampSupValeur('NUMERO',0);
  VH_GC.TOBRGPRE.AddChampSupValeur('INDICEG',0);
  VH_GC.TOBRGPRE.AddChampSupValeur('DATEMODIF',iDate1900);
end;

Procedure InitLaVariableGC ;
  function GetSqlDevise: string;
  begin
    Result := 'SELECT D_DEVISE, D_PARITEEUROFIXING, D_DECIMALE, D_ARRONDIPRIXACHAT, D_ARRONDIPRIXVENTE'
            + ' FROM DEVISE'
  end;
  function GetSqlGAR: string;
  begin
    Result := 'SELECT *'
            + ' FROM ARRONDI'
            + ' ORDER BY GAR_CODEARRONDI,GAR_VALEURSEUIL DESC'
  end;

//GP_20080616_MM_GP14941 Déb
  function GetSqlYFO: string;
  begin
    Result := 'SELECT * FROM YTARIFSPARAMETRES'
  end;
//GP_20080616_MM_GP14941 Fin

//GP_20080625_DKZ_TD9702 Déb
{$IFDEF GCGC}
  function GetSqlAJF: string;
  begin
    Result := 'SELECT IIF((AJF_JOURFIXE="-"), AJF_ANNEE, 1900) AS IANNEE'
             +     ', AJF_MOIS AS IMOIS, AJF_JOUR AS IJOUR'
             +     ' FROM JOURFERIE'
             +     ' ORDER BY IANNEE DESC, IMOIS, IJOUR';
  end;
{$ENDIF GCGC}

//GP_20080625_DKZ_TD9702 Fin
BEGIN
{$IFNDEF EAGLSERVER}
  VH_GC:=LaVariableGC.Create ;
{$ENDIF !EAGLSERVER}
	TheParamTva := TTVSTA2014.Create;
  TheParamTva.assigned := false;
  
VH_GC.TobNumGroupe := Tob.Create ('', nil, -1);
VH_GC.TOBGCE:=TOB.Create('',Nil,-1) ;
VH_GC.TOBGCB:=TOB.Create('',Nil,-1) ;
VH_GC.TOBGCBDGD:=TOB.Create('',Nil,-1) ;
//
VH_GC.TOBRGDOC :=TOB.Create('',Nil,-1) ;
InitChampsSupSTORG;
VH_GC.TOBRGPRE :=TOB.Create('',Nil,-1) ;
InitChampsSupRGPRE;
//
VH_GC.TOBPieceEDT := TOB.Create ('PIECE',nil,-1);
//
VH_GC.TOBEdt:=TOB.Create('',Nil,-1) ;
VH_GC.TOBPCaisse:=TOB.Create('PARCAISSE',Nil,-1) ; VH_GC.TOBPCaisse.InitValeurs ;
{$IFDEF CHR}
VH_GC.GCTOBEtab := TOB.Create('',Nil,-1);
{$ENDIF CHR}
VH_GC.AFTOBTraduction:=TOB.Create('',Nil,-1) ;
{$ifdef GIGI}
VH_GC.AFTOBCatalogue:=TOB.Create('',Nil,-1) ;
{$endif}
VH_GC.BOTypeMasque_Defaut:='...' ; // Utilisé uniquement par Mode

VH_GC.TobDevise := tMemoryTob.Create('_DEVISE_', GetSqlDevise);

{$IFDEF BTP}

VH_GC.BTGestParag := true;
VH_GC.BTPrixMarche := true;
VH_GC.TOBArtEcart := TOB.Create ('ARTICLE',nil,-1);
VH_GC.TOBParamTaxe := TOB.Create ('LES PARAMTAXE',nil,-1);
VH_GC.TOBBPM := TOB.Create ('LES PIECES MILLIEME',nil,-1);
VH_GC.TOBTABFERIE := TOB.Create ('LE PARAM JF',nil,-1);
VH_GC.BSVUploadOK := false;
VH_GC.BSVOpenDoc := false;

{$ELSE}
VH_GC.BTGestParag := false;
VH_GC.BTPrixMarche := false;
{$ENDIF}
VH_GC.IsChpsOblig := False; // JTR test champs obligatoire pièce
VH_GC.ModeGestionEcartComptable := ''; {DBR CPA}
{$IFDEF NOMADE}
InitLaVariableGCPCP;
{$ENDIF NOMADE}

//Pour différencier Business Suite et Place
//forcer à True pour toutes les gescoms mais initialisé dans UtilDispGC
// pour la Gestion Commerciale PGE en fonction de la seria.
{$IFNDEF CRM}
VH_GC.GCAchatStockSeria:=True ;
{$ENDIF !CRM}
{$IFDEF CHR}
VH_GC.LeSiteCourant := '';  // LB - Journal des évènements
{$ENDIF CHR}

// Si Module PAIE Non sérialisé et DA dans paramsoc on charge les paramètres de la paie
{$IFDEF GCGC} //mcd 04/12/06 passe de Gescom en GCGC
{$IFNDEF PAIEGRH}
{$IFNDEF OGC}
{$IFNDEF PGIMAJVER}
{$IFNDEF CRM}
{$IFNDEF CMPGIS35}
InitLaVariablePaie; // CCMX-CEGID ORGANIGRAMME DA
{$ENDIF}
{$ENDIF !CRM}
{$ENDIF !PGIMAJVER}
{$ENDIF OGC}
{$ENDIF PAIEGRH}
  { Gestion de WPARAMFONCTION en mémoire }
  VH_GC.ParamFonction := TParamFonction.Create();
  VH_GC.ParamFonction.Initialisation();
{$ENDIF GCGC}
  {$IFDEF AFFAIRE}
  VH_GC.VLibellesDesAffaires := TAFO_LibelleAffaire.Create; // BDU - 02/02/07
  VH_GC.VTiersParc := TAFO_TiersParc.Create;  // BDU - 02/04/07
  VH_GC.AfTobBlocageAff := nil; // ST  07/11/07
  {$ENDIF}
//GP_20080616_MM_GP14941
  VH_GC.TobYFO := tMemoryTob.Create('YFO', GetSqlYFO);
//GP_20080625_DKZ_TD9702 Déb
{$IFDEF GCGC}
  VH_GC.TobAJF := tMemoryTob.Create('AJF', GetSqlAJF); { Tob des jours fériés }
{$ENDIF GCGC}                                          
END ;

{$IFDEF NOMADE}
Procedure InitLaVariableGCPCP;
begin
  // DCA 01/09/04 - Initialisation déplacé dans UtilDispGc.pas AfterProtectGc
  //VH_GC.PCPVenteSeria := True;
  //VH_GC.PCPAchatSeria := True;
  VH_GC.PCPPrefixe := '';
  VH_GC.PCPRepresentant := '';
  VH_GC.PCPPceVente := '';
  VH_GC.PCPPceAchat := '';
  VH_GC.PCPUsVte := false;
  VH_GC.PCPUsAch := false;
end;
{$ENDIF}

Procedure LibereLaVariableGC ;
Begin
  {$IFNDEF EAGLSERVER}
  if Assigned(VH_GC) then
  {$ELSE}	// CHR  - ajouté sinon violation d'accès en precosses serveur
  if VH_GC <> nil then
  {$ENDIF EAGLSERVER}
  begin
    FreeAndNil (VH_GC.TobNumGroupe);
    {$IFDEF AFFAIRE}
      FreeAndNil(VH_GC.VLibellesDesAffaires); // BDU - 02/02/07
      FreeAndNil(VH_GC.VTiersParc); // BDU - 02/04/07
    {$ENDIF AFFAIRE}
{$IFDEF BTP}
	 VH_GC.TOBParamTaxe.free;
   VH_GC.TOBArtEcart.Free;
	 VH_GC.TOBBPM.free;
   VH_GC.TOBTABFERIE.free;
{$ENDIF}
    VH_GC.MTOBMEA.Free;
    VH_GC.TOBParPiece.Free;
    VH_GC.MTOBParPieceComp.Free;
    VH_GC.MTOBParPieceDomaine.Free;
    VH_GC.TOBGCE.Free;
    VH_GC.TOBGCB.Free;
    VH_GC.TOBGCBDGD.Free;
    VH_GC.TOBRGDOC.free;
    VH_GC.TOBRGPRE.free;
    VH_GC.TOBPieceEDT.free;
    VH_GC.TOBEdt.Free;
    FreeAndNil(VH_GC.TobDevise);
    VH_GC.MGCTOBDim.Free;
    VH_GC.TOBPCaisse.Free;
    VH_GC.MGCTOBAna.Free;
    VH_GC.AFTOBTraduction.Free;
    VH_GC.MTobChpsOblig.Free;
    {$IFDEF GIGI}
       VH_GC.AFTOBCatalogue.Free;
    {$ENDIF GIGI}
    {$IFDEF AFFAIRE}
      VH_GC.AFORessourceUser.Free;
    {$ENDIF AFFAIRE}
    VH_GC.AFTOBAffectChmp.free;
    VH_GC.MTobChpsOblFic.Free;
    VH_GC.MTobFicConf.Free;
    VH_GC.MTobRestrFiche.Free;
    //CHR _20071220_GFO
    {$IFDEF CHR}
    FreeAndNil(VH_GC.GCTOBEtab);
    {$ENDIF CHR}
    FreeAndNil(VH_GC.TobWWF);
//GP_20080616_MM_GP14941
    FreeAndNil(VH_GC.TobYFO);
    {$IFDEF STK}
      with VH_GC Do
      begin
        FreeAndNil(TobSDI);
        FreeAndNil(TobSFL);
        FreeAndNil(TobGC1);
        FreeAndNil(TobGSN);
        FreeAndNil(TobGVP);
        FreeAndNil(TobGVT);
        FreeAndNil(TobGSR);
        FreeAndNil(TobTemp);
        FreeAndNil(TobWWA);
      end;
      {$IF Defined(GCGC) and not Defined(PGIMAJVER)}
        ConversionQteGA('', 1, '', ''); { Pour libérer LastToBGA }
      {$IFEND GCGC && !PGIMAJVER}
    {$ENDIF STK}

    if assigned(VH_GC.TOBTiersImpactPce) then FreeAndNil(VH_GC.TOBTiersImpactPce);
    if assigned(VH_GC.TOBElementImpactPce) then FreeAndNil(VH_GC.TOBElementImpactPce);
    if assigned(VH_GC.RecupInfos) then
    begin
      VH_GC.RecupInfos.Clear;
      FreeAndNil(VH_GC.RecupInfos);
    end;
    if assigned(VH_GC.TobTauxNegSoc) then FreeAndNil(VH_GC.TobTauxNegSoc);
    if assigned(VH_GC.TobDupPceTiers) then FreeAndNil(VH_GC.TobDupPceTiers);
    if assigned(VH_GC.TobElementCptaPce) then FreeAndNil(VH_GC.TobElementCptaPce);
    if assigned(VH_GC.TobTypeCptaException) then FreeAndNil(VH_GC.TobTypeCptaException);
    {$IFDEF GCGC}
      { Gestion de WPARAMFONCTION en mémoire }
      if Assigned(VH_GC.ParamFonction) then
        FreeAndNil(VH_GC.ParamFonction);
//GP_20080625_DKZ_TD9702 Déb
      if Assigned(VH_GC.TobAJF) then
        FreeAndNil(VH_GC.TobAJF);
//GP_20080625_DKZ_TD9702 Fin
    {$ENDIF GCGC}
    {$IFDEF AFFAIRE}
    if assigned(VH_GC.AfTobBlocageAff) then FreeAndNil(VH_GC.AfTobBlocageAff); // ST  07/11/07
    {$ENDIF AFFAIRE}
    {$IFDEF HRPGI}
    if LookUpCurrentSession.UserObjects.IndexOf('VH_GC') >= 0 then
      LookUpCurrentSession.UserObjects.Delete(LookUpCurrentSession.UserObjects.IndexOf('VH_GC')) ;
    {$ENDIF}
    {$IFNDEF EAGLSERVER}
      VH_GC.Free ; VH_GC:=Nil ;
    {$ENDIF !EAGLSERVER}
  end;
  // CCMX-CEGID ORGANIGRAMME DA
  {$IFDEF GCGC} // BDU : Remplace GESCOM par GCGC
	  {$IFNDEF EAGLSERVER}
	    {$IFNDEF PAIEGRH}
        {$IFNDEF PGIMAJVER}
          {$IFNDEF CRM}
{$IFNDEF CMPGIS35}
            If VH_Paie <> Nil Then
            Begin
              VH_Paie.Free;
              VH_Paie:=Nil;
            End;
{$ENDIF}
          {$ENDIF}
        {$ENDIF !PGIMAJVER}
	    {$ENDIF !PAIEGRH}
	  {$ENDIF !EAGLSERVER}
  {$ENDIF GESCOM}
  // FIN CCMX-CEGID ORGANIGRAMME DA
  if TheParamTva <> nil then TheParamTva.Free;
End;

Function OnChangeUserGC : boolean;
Begin
{$IFNDEF BTP}
{$IFDEF AFFAIRE}
  // on libère l'ancienne ressource associée au user courant
  if (VH_GC.AFORessourceUser <> nil) then
  	begin
    VH_GC.AFORessourceUser.Free;
    VH_GC.RessourceUser := ''; // Il faut absolument vider l'ancienne ressource pour recharger la suivante
  	end;

  // on va rechercher la nouvelle ressource associée au user courant
  VH_GC.AFORessourceUser := TAFO_Ressources.Create;
  if (VH_GC.AFORessourceUser.AddRessourceD1User(V_PGI.User, false, true, true) = -1) then
		// On n'a pas de ressource associée au user courant
		begin
		VH_GC.AFORessourceUser.Free;
		VH_GC.AFORessourceUser := nil;
		end
	else
    // affectation du code ressource associé au user courant
    //VH_GC.RessourceUser:=GetRessourceUser;
    VH_GC.RessourceUser:=VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetString('ARS_RESSOURCE');

  VH_GC.UserInvite := TraiteUserinvite;
{$ENDIF}
{$ENDIF}
Result:=True;
End;

Procedure ChargeTobsConfid ;
var requete : String;
begin
// TobFicConf: élements utiles tablette YYPARAMOBLIG
// MB Ajout de tout les champs pour être compatible avec la replication.
requete := 'Select CO_TYPE, CO_CODE, CO_LIBELLE, ' +
            'CO_ABREGE, CO_LIBRE from COMMUN where CO_TYPE="POB"' ;

if (ctxaffaire in V_PGI.PGIContexte)  then
   begin
  // gm 28/12/04 qd le complément ligne sera ok , on mettra LIKE "GP%"
  if VH_GC.GASeria then requete  :=requete + 'AND (CO_CODE like "A%" OR CO_CODE like "B%" OR CO_CODE ="GP1" OR CO_CODE ="GP3" '
      else requete :=requete +'AND (CO_CODE = "A09" OR CO_CODE = "B01" OR CO_CODE ="B02" OR CO_CODE ="B06" '; // cas GI pas sérialisé, uniquement GRC
  end
else
  if (ctxGCAFF in V_PGI.PGIContexte)  then
   requete :=requete +' AND ( CO_CODE like "G%" or CO_CODE like "A%" '
  else
    if (ctxGPAO in V_PGI.PGIContexte)  then
      requete := requete +' AND ( CO_CODE like "G%" or CO_CODE like "W%" '
    else
      requete :=requete +' AND ( CO_CODE like "G%" ';

{$IFDEF GIGI}   
  if ctxGRC in V_PGI.PGIContexte then requete  :=requete  +' OR CO_CODE LIKE "X%"';
{$ENDIF}
if ctxGRC in V_PGI.PGIContexte then requete :=requete + ' OR CO_CODE like "R%")'
else
{$IFDEF GIGI} //mcd 04/12/07 en GI tiers 360 vu même si pas GRC
  requete  :=requete  +' OR CO_CODE = "R10")';
{$else}
 requete :=requete + ')';
{$ENDIF}

  if (ctxscot in V_PGI.PGIContexte) then //  gm 26/5/5 pas d'article contrat pour scot
    requete :=requete + ' AND (CO_CODE <> "A19")';

{$IFDEF CHR}
requete := 'Select * from COMMUN where CO_TYPE="POB"';
{$ENDIF}

  if VH_GC.MTObFicConf <>Nil then VH_GC.MTObFicConf.free; //mcd 16/03/2005 si réappel fct  ..erreur memcheck

// TObFicConf : ajout des spécifiques GEP
Requete := Requete + ' union ' +
            'Select YX_TYPE as CO_TYPE, YX_CODE as CO_CODE, YX_LIBELLE as CO_LIBELLE, ' +
            'YX_ABREGE as CO_ABREGE, YX_LIBRE as CO_LIBRE from CHOIXEXT where YX_TYPE="CFD"';

  VH_GC.MTobFicConf := TMemoryTob.Create ('CO_POB', GetSqlCO_POB (Requete));


// TobChpsOblFic : liste des champs obligatoires
  if VH_GC.MTObChpsOblFic <>Nil then VH_GC.MTObChpsOblFic.free; //mcd 16/03/2005 si réappel fct
  VH_GC.MTobChpsOblFic := TMemoryTob.Create ('GOB', GetSqlGOB_Champ);


//TobRestrFiche
requete := 'Select * from PARAMOBLIG where GOB_GRVISIBLE<>"" or GOB_GRENABLE<>"" or GOB_USVISIBLE1<>"" or GOB_USENABLE1<>""';
  if VH_GC.MTObRestrFiche <>Nil then VH_GC.MTObRestrFiche.free; //mcd 16/03/2005 si réappel fct
  VH_GC.MTobRestrFiche := TMemoryTob.Create ('GOB', GetSqlGOB_Fiche);

  {$IFDEF BTP}
  VH_GC.MTobChpsOblFic.load;
  VH_GC.MTobFicConf.load;
  VH_GC.MTobRestrFiche.load;
  {$ENDIF}
end;
{$ENDIF}  //GCGC

function GetNaturePieceCde (ForSql : boolean=true) : string;
begin
{$IFDEF BTP}
if ForSql then result := 'CC","CBT","PRE'
          else result := 'CC;CBT;PRE;';
{$ELSE}
if ForSql then result := 'CC","PRE'
          else result := 'CC;PRE;'
{$ENDIF}
end;

function IsLivraisonClient (TOBL : TOB) : Boolean;
begin
	result := false;
  if TOBL.NomTable = 'LIGNE' then
  begin
  	if Pos(TOBL.GetValue('GL_NATUREPIECEG'),GetNaturePieceBLC(false)) > 0 then result := true;
  end else if TOBL.nomTable = 'PIECE' then
  begin
  	if Pos(TOBL.GetValue('GP_NATUREPIECEG'),GetNaturePieceBLC(false)) > 0 then result := true;
  end;
end;

function IsLivraisonClientBTP (TOBL : TOB) : Boolean;
begin
	result := false;
  if TOBL.NomTable = 'LIGNE' then
  begin
  	if Pos(TOBL.GetValue('GL_NATUREPIECEG'),GetNaturePieceLBT(false)) > 0 then result := true;
  end else if TOBL.nomTable = 'PIECE' then
  begin
  	if Pos(TOBL.GetValue('GP_NATUREPIECEG'),GetNaturePieceLBT(false)) > 0 then result := true;
  end;
end;

function IsLivraisonClientNEG (TOBL : TOB) : Boolean;
begin
	result := false;
  if TOBL.NomTable = 'LIGNE' then
  begin
  	if Pos(TOBL.GetValue('GL_NATUREPIECEG'),GetNaturePieceNEG(false)) > 0 then result := true;
  end else if TOBL.nomTable = 'PIECE' then
  begin
  	if Pos(TOBL.GetValue('GP_NATUREPIECEG'),GetNaturePieceNEG(false)) > 0 then result := true;
  end;
end;

function GetNaturePieceNEG (ForSql : boolean=true) : string;
begin
if ForSql then result := 'BLC'
          else result := 'BLC;';
end;

function GetNaturePieceLBT (ForSql : boolean=true) : string;
begin
if ForSql then result := 'LBT'
          else result := 'LBT;';
end;

function GetNaturePieceBLC (ForSql : boolean=true) : string;
begin
{$IFDEF BTP}
if ForSql then result := 'BLC","LBT'
          else result := 'BLC;LBT;';
{$ELSE}
if ForSql then result := 'BLC'
          else result := 'BLC;'
{$ENDIF}
end;


Function GetPieceAchat (WithRetour: boolean = True;WithCF : boolean=false;ForSql : boolean=true; WithRecepRegroupe : boolean=true;WithpropositionAchat : boolean=false): string;
begin
  if ForSql then
  begin
    Result := '"BLF","FF"';
    if WithCF then Result := Result + ',"CF","CFR"';
    if WithRetour then Result := Result + ',"BFA"';
    if WithRecepRegroupe then Result := Result + ',"LFR"';
    if WithPropositionAchat then Result := Result + ',"DEF"';
  end else
  begin
    Result := 'BLF;FF';
    if WithCF then Result := Result + ';CF;CFR';
    if WithRetour then Result := Result + ';BFA';
    if WithRecepRegroupe then Result := Result + ';LFR';
    if WithPropositionAchat then Result := Result + ';DEF';
    result := Result + ';';
  end;
end;


{$IFDEF GPAO}
procedure GPCreationUNIDansMEA;
var
  TobUNI: Tob;
begin
  if Assigned(VH_GC.MTOBMea) then
  begin
    VH_GC.MTobMEA.Load;
    TobUni := VH_GC.MTOBMea.FindFirst(['GME_QUALIFMESURE','GME_MESURE'], ['PIE','UNI'], False);
    if TobUni = nil then
    begin
      TobUNI := Tob.Create('MEA', nil, -1);
      try
        TobUni.SetString('GME_QUALIFMESURE', 'PIE');
        TobUni.SetString('GME_MESURE', 'UNI');
        TobUni.SetString('GME_LIBELLE', TraduireMemoire('Unité'));
        TobUni.SetDouble('GME_QUOTITE', 1);
        TobUni.InsertOrUpdateDB(False);
      finally
        TobUNI.Free;
      end;
    end;
  end;
end;
{$ENDIF GPAO}

{$IFDEF GPAO}
procedure GPVerificationQuotiteHHCC;
var
  TobUNI: Tob;
begin
  if Assigned(VH_GC.MTOBMea) then
  begin
    VH_GC.MTobMEA.Load;
    TobUni := VH_GC.MTOBMea.FindFirst(['GME_QUALIFMESURE','GME_MESURE'], ['THC','HC'], False);
    if TobUni = nil then
    begin
      TobUNI := Tob.Create('MEA', nil, -1);
      try
        TobUni.SetString('GME_QUALIFMESURE', 'THC');
        TobUni.SetString('GME_MESURE', 'HC');
        TobUni.SetString('GME_LIBELLE', TraduireMemoire('Heure centième'));
        TobUni.SetDouble('GME_QUOTITE', 3600);
        TobUni.InsertOrUpdateDB(False);
      finally
        TobUNI.Free;
      end;
    end
    else
    begin
      TobUni.SetDouble('GME_QUOTITE', 3600);
      TobUni.UpdateDB(False);
    end;
  end;
end;
{$ENDIF GPAO}

//mcd 10/03/20005 : fct faire une fois à la connection, au lieu de le faire à chaque utilisation ...
//Penser de modifier à l'identique si modif de GereCritGroupeConf....
// dans GalOutils du DP
function GIGereCritGroupeConf : string;
Var Q: TQuery;
    critere, plus : String;
begin
  // si superviseur, le crit. restera vide car on voit tout
  // 27/07/01 sauf en Mode pme : cacher les dossiers qui ont une conf.
  critere := '';
  // pour ne sortir aucuns groupes sauf <<Tous>>
  plus := 'GRP_CODE=""';
  // 1er critère = les dossiers sans confidentialité...
  critere := '';
  // dans la combo DOS_GROUPECONF, on ne met
  // que les groupes auxquels appartient le user

  // MD 07/11/2007 - FQ 11811 - élimination des vues
  {
  Q := OpenSQL('SELECT UCO_GROUPECONF FROM ##DP##.USERCONF WHERE UCO_USER="'+V_PGI.User+'"',True,-1,'',True);
  try
    While Not Q.Eof do
    begin
      plus := plus + ' OR GRP_CODE="'+ Q.FindField('UCO_GROUPECONF').AsString +'"';
      critere := critere + Q.FindField('UCO_GROUPECONF').AsString+ ';' ;  //mcd 19/09/03 ordre des champs changé (mise XXX; au liue de ;XXX plantage guillement droit manquant}

  Q := OpenSQL('SELECT GRP_CODE FROM GRPDONNEES, LIENDONNEES WHERE GRP_NOM = LND_NOM AND GRP_NOM = "GROUPECONF" AND GRP_ID= LND_GRPID '
              +'AND LND_USERID="'+V_PGI.User+'"',True,-1,'',True);
  try
    While Not Q.Eof do
    begin
      plus := plus + ' OR GRP_CODE="'+ Q.FindField('GRP_CODE').AsString +'"';
      critere := critere + Q.FindField('GRP_CODE').AsString+ ';' ;
      // Next OK, on doit lire tous les groupes auxquels appartient le user
      Q.Next ;
    end;
  finally
    Ferme(Q);
  end;
  //On retourne le critère du ThValMultiComboBox et sa variable "Plus" .
  Result := critere+';ComboPlus;'+Plus;
end;

{$ifdef GCGC}
//mcd 11/03/2005
//fct qui permet de charger toutes les équipes affectées à la ressource
// qui s'est connecté.
// n'est fait que si gestion confidentialité sur planning
function GetAfEquipe : string;
Var Q: TQuery;
begin
  Result := '';
  if GetParamsoc ('SO_AFTypeConPla') <> 'EQU' then exit;
  Q := OpenSQL('SELECT REE_EQUIPERESS FROM RESSOURCEEQUIPE WHERE REE_RESSOURCE="'+VH_GC.RessourceUser+'"',True,-1,'',True);
  try
    While Not Q.Eof do
    begin
      Result := Result + Q.FindField('REE_EQUIPERESS').AsString+ ';' ;
      // Next OK, on doit lire tous les equipes associées à la ressource
      Q.Next ;
    end;
  finally
    Ferme(Q);
  end;
  if Result = '' then
    begin  //on va chercher l'équipe de la ressource
   (* Q := OpenSQL('SELECT ARS_EQUIPERESS FROM RESSOURCE WHERE ARS_RESSOURCE="'+VH_GC.RessourceUser+'"',True);
    If Not Q.Eof then Result := Result + Q.FindField('ARS_EQUIPERESS').AsString+ ';' ; *)
			//mcd 13/02/2006 on ne refait aps la requête
{$IFNDEF BTP}
{$ifdef AFFAIRE}
		if (VH_GC.AFORessourceUser <> nil) then // Ajout PL le 13/02/06
      if VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetString('ARS_EQUIPERESS') <>'' then
          Result := Result + VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetString('ARS_EQUIPERESS')+ ';' ;
 {$endif}
{$ENDIF}
   if Result=';' then result:='';
(*    Ferme(Q); *)
    end;
end;
{$ENDIF GCGC}

{$IFDEF GCGC}
function GetParamSocGcGestUniteMode: String;
begin
  Result := GetParamSocSecur(SO_gcGestUniteMode, '');
end;
{$ENDIF GCGC}

{$IFDEF GCGC}
function GetParamSocGCDefQualifUniteGA: String;
begin
  Result := GetParamSocSecur('SO_GCDEFQUALIFUNITEGA', '');
end;
{$ENDIF GCGC}

{$IFDEF GCGC}
function GetParamSocGCDefUniteGA: String;
begin
  Result := GetParamSocSecur('SO_GCDEFUNITEGA', '');
end;
{$ENDIF GCGC}

{$IFDEF STK}
procedure InitVH_GCGestUnitMode;
var
  Code: String;
begin
  Code := GetParamSocGcGestUniteMode;
  if Code = '003' then
    VH_GC.GestUniteMode := meaAdvConvU
  else if Code = '002' then
    VH_GC.GestUniteMode := meaAdvanced
  else if Code = '001' then
    VH_GC.GestUniteMode := meaNormal
  else
  begin
    PGIError(Format(TraduireMemoire('Le paramètre société pilotant la gestion des unités a une valeur incorrecte ! %s Merci de contacter le SAT. %s %s Pour cette session, la gestion des unités est initialisée en mode ''normal''.'), [Code, #13, #13, #13]));
    VH_GC.GestUniteMode := meaNormal;
    Code := '001';
  end;

  if IsDayPass then
  begin
    Debug('---------------------------------------------');
    Debug(Format(TraduireMemoire('Gestion des unités : %s '), [Code]));
    Debug('---------------------------------------------');
  end;

  if GetParamSocGCDefQualifUniteGA = '' then
    SetParamSoc('SO_GCDEFQUALIFUNITEGA', 'PIE');
  if GetParamSocGCDefUniteGA = '' then
    SetParamSoc('SO_GCDEFUNITEGA', 'UNI');
end;
{$ENDIF STK}

{$IFDEF GCGC}
function GetParamSocGCNewLienPiece: Boolean;
begin
  Result := GetParamSocSecur('SO_GCNEWLIENPIECE', False);
end;

{$IFDEF CHR}
{$IFDEF EAGLCLIENT}
function GetInfoGCTOBEtab( Etablissement, NomChamp : string ) : string;
var
  TOBEtab : TOB;
begin
  Result := '';
  TOBEtab := TOBRechercheOptimise( VH_GC.GCTOBEtab, 'ET_ETABLISSEMENT', Etablissement );
  if TOBEtab <> nil then
    Result := TOBEtab.GetString( NomChamp );
end;
{$ENDIF EAGLCLIENT}

{***********A.G.L.***********************************************
Auteur  ...... : Lydie Barbey
Créé le ...... : 01/09/2005
Modifié le ... : 11/10/2005
Description .. : Initialisation de la variable VH_GC.LesSiteCourant avec le
Suite ........ : site courant configurer dans la TOX
Mots clefs ... :
*****************************************************************}
procedure ChargeVHGCLeSiteCourant;
begin

  VH_GC.LeSiteCourant := GetColonneSQL('STOXSITES','##TOP 1## SSI_CODESITE', 'SSI_CURRENTSITE="X" AND SSI_SITEENABLED="X"');
  if VH_GC.LeSiteCourant='' then
    VH_GC.LeSiteCourant := 'SC';
end;
{$ENDIF CHR}

{$ENDIF GCGC}

function JAiLeDroitMenu122: Boolean;
begin
  Result := JAiLeDroitTag(122100) or JAiLeDroitTag(122200) or JAiLeDroitTag(122300) or JAiLeDroitTag(122400)
end;

Function GereSTAGC: boolean;
begin
  Result := (GetParamSoc('SO_ASSEMBLAGE')) and (GetParamSocSecur('SO_STAGC', False));
end;

{***********A.G.L.***********************************************
Auteur  ...... : PCH - Philippe Chevron
Créé le ...... : 10/02/2006
Modifié le ... :   /  /
Description .. : Dans le cas de personnalisation des paramètres société,
Suite ........ : on va surcharger les paramètres d'origine par ceux qui ont
Suite ........ : été personnalisés.
Mots clefs ... : PutParamTob

  ATTENTION :
  La fonction PutParamTob utilisée pour surcharger les paramètres société dans la
  V_PGI.TOBSOC ne fonctionne que si le paramètre société existe déjà dans V_PGI.TOBSOC.
  A l'entrée dans l'application, dans Ent1, CHARGEMAGHALLEY => CHARGESOCIETEHALLEY
  => ChargeInfosSociete fait le tout 1er GetParamSocSecur sur 'SO_SOCIETE'.
  A cette occasion est fait un ChargeParamSoc qui crée V_PGI.TOBSOC avec
  tous les paramètres société.
  (SELECT SOC_NOM, SOC_DATA, SOC_DESIGN FROM PARAMSOC WHERE SOC_NOM LIKE "SO_%")
  Si un jour, le fonctionnement initial était modifié et que tous les paramètres
  société n'étaient pas créés d'office, il faudrait faire un GetParamSocSecur sur
  le paramètre à surcharger avant l'appel à PutParamTob.

*****************************************************************}
{$IFDEF GIGI}
Procedure ChargeParamsocPerso;
{$IFNDEF EAGLSERVER}
Var Q : TQuery;
    St, StValParam : String;
    StEtablissement, StDepartement : String;
    Table_Libre : Array[1..10] Of String ;
    Table_ParamPerso : Array[1..3] Of String ;
    TOBParamSocPerso : TOB;
    I_i1, I_i2 : Integer;
{$endiF EAGLSERVER}
BEGIN

{$IFNDEF EAGLSERVER}
  if (GetParamSocSecur('SO_AFPERSOCHOIX',False)) then
  begin
    // Une personnalisation des paramètres société a été effectuée

    // Détermination de la personnalisation des paramètres société
    Table_ParamPerso[1] := GetParamSocSecur('SO_AFPERSOPARAM1','');
    Table_ParamPerso[2] := GetParamSocSecur('SO_AFPERSOPARAM2','');
    Table_ParamPerso[3] := GetParamSocSecur('SO_AFPERSOPARAM3','');

    // L'initialisation du pointeur sur la ressource liée à l'utilisateur courant
    // est faite dans OnChangeUserGC ainsi que la lecture de la ressource
    // Les infos de la ressource sont dans le 1er élément de la propriété GetRessource
    // de l'objet
    // La désallocation se fait dans LibereLaVariableGC

    StEtablissement := ''; // Ajout PL le 13/02/06
    StDepartement := '';
      for I_i1 := 1 to 10 do
      begin
        Table_Libre[I_i1] := '';
      end;

		if (VH_GC.AFORessourceUser <> nil) then // Ajout PL le 13/02/06
			begin
      StEtablissement := VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetValue('ARS_ETABLISSEMENT');
      StDepartement := VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetValue('ARS_DEPARTEMENT');
      for I_i1 := 1 to 9 do
      begin
        Table_Libre[I_i1] := VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetValue('ARS_LIBRERES'+IntToSTR(I_i1));
      end;
      Table_Libre[10] := VH_GC.AFORessourceUser.GetRessource(0).Tob_Champs.GetValue('ARS_LIBRERESA');
			end;


    for I_i1 := 1 to 3 do
    begin
      if (Trim(Table_ParamPerso[I_i1]) = '')
      or (Trim(Table_ParamPerso[I_i1]) = 'AB_') then Table_ParamPerso[I_i1] := ''
      else if Table_ParamPerso[I_i1] = 'AB0' then Table_ParamPerso[I_i1] := StEtablissement
      else if Table_ParamPerso[I_i1] = 'ABB' then Table_ParamPerso[I_i1] := StDepartement
      else if Table_ParamPerso[I_i1] = 'ABA' then Table_ParamPerso[I_i1] := Table_Libre[10]
      else
      begin
        for I_i2 := 1 to 9 do
        begin
          if Table_ParamPerso[I_i1] = 'AB' + IntToSTR(I_i2) then Table_ParamPerso[I_i1] := Table_Libre[I_i2];
        end;
      end;
    end;
    TOBParamSocPerso := Tob.Create('les paramsoc personnalises', nil, -1);
    try
      // Recherche des paramètres société personnalisés liés à l'utilisateur
      St := 'SELECT * FROM AFPARAMSOC WHERE APA_PARAM1 = "' + Table_ParamPerso[1] + '"'
                                    + ' AND APA_PARAM2 = "' + Table_ParamPerso[2] + '"'
                                    + ' AND APA_PARAM3 = "' + Table_ParamPerso[3] + '"';
      Q := OpenSQL(St, TRUE, -1, 'AFPARAMSOC',True) ;
      try
        TOBParamSocPerso.LoadDetailDB('AFPARAMSOC','','',Q,FALSE,FALSE) ;
        // Surcharge des paramètres société par ceux personnalisés liés à l'utilisateur
        I_i2 := 0;
        while I_i2 < TOBParamSocPerso.Detail.count do
        begin
          if TOBParamSocPerso.Detail[I_i2].GetValue('APA_PERSOTYPECHAMP') = 'B' then
            //CheckBox : transformer X en -1 et - en 0
            if TOBParamSocPerso.Detail[I_i2].GetValue('APA_DATA') = 'X' then
              StValParam := '-1'
            else
              StValParam := '0'
          else
            StValParam := TOBParamSocPerso.Detail[I_i2].GetValue('APA_DATA');

          PutParamTob(TOBParamSocPerso.Detail[I_i2].GetValue('APA_NOMPARAMSOC'), StValParam);
          inc(I_i2);
        end;
      finally
        Ferme(Q) ;
      end;
    finally
      TOBParamSocPerso.free;
    end;
  end;
{$ENDIF EAGLSERVER}
END;
{$ENDIF GIGI}

{$IFDEF GCGC}
procedure LaVariableGC.Initialisation;
begin
  {$IFDEF STK }
    if TobGSN <> nil then TobGSN.ClearDetail;
    if TobSDI <> nil then TobSDI.ClearDetail;
    if TobGC1 <> nil then TobGC1.ClearDetail;
    if TobSFL <> nil then TobSFL.ClearDetail;
    if TobGSR <> nil then TobGSR.ClearDetail;
    if TobGVP <> nil then TobGVP.ClearDetail;
    if TobGVT <> nil then TobGVT.ClearDetail;
  {$ENDIF STK}
  if TobDevise <> nil then TobDevise.ClearDetail;
//GP_20080616_MM_GP14941
  if TobYFO    <> nil then TobYFO.ClearDetail;
//GP_20080625_DKZ_TD9702
  if TobAJF    <> nil then TobAJF.ClearDetail;
end;
{$ENDIF GCGC}

{$IFDEF STK}
function GetSqlGSN: string;
begin
  Result := 'SELECT GSN_QUALIFMVT,GSN_LIBELLE,GSN_STKTYPEMVT,GSN_QTEPLUS'
          + ',GSN_CONTREMARQUE'
          + ',GSN_QUALIFMVTSUIV,GSN_SIGNEMVT,GSN_STKFLUX,GSN_MAJPRIXVALO,GSN_CALLGSL,GSN_CALLGSS'
          + ',GSN_SDISPODISPATCH,GSN_SFLUXDISPATCH,GSN_GERESERIEGRP'
          + ',IIF(SUBSTRING(GSN_SDISPOPICKING,1,1)="<", "", GSN_SDISPOPICKING) AS GSN_SDISPOPICKING'
          + ',IIF(SUBSTRING(GSN_SFLUXPICKING,1,1)="<", "", GSN_SFLUXPICKING) AS GSN_SFLUXPICKING'
          + ',ISNULL(GSN_CTRLDISPO,"000") AS GSN_CTRLDISPO,ISNULL(GSN_CTRLPEREMPTION,"000") AS GSN_CTRLPEREMPTION'
          + ' FROM STKNATURE'
          + ' ORDER BY GSN_QUALIFMVT'
end;
{$ENDIF STK}

{$IFDEF STK}
function GetSqlSDI: string;
begin
  Result := 'SELECT CO_CODE AS SDI_STATUTDISPO'
          + ',SUBSTRING(CO_ABREGE, 3, 14) AS SDI_COMPTEUR, Trim("#ICO#"||STR(CO_LIBRE)) AS SDI_IMAGE'
          + ' FROM COMMUN'
          + ' WHERE CO_TYPE="SDI"'
          + ' ORDER BY CO_CODE'
end;
{$ENDIF STK}

{$IFDEF STK}
function GetSqlGC1: string;
begin
  Result := 'SELECT CO_LIBRE AS GC1_COMPTEUR'
          + ',CO_LIBELLE AS GC1_LIBELLE'
          + ',CO_ABREGE AS GC1_STATUTDISPO'
          + ' FROM COMMUN'
          + ' WHERE CO_TYPE="GC1"'
          + ' ORDER BY CO_LIBRE'
end;
{$ENDIF STK}

function GetSqlGCETT : string;
begin
  Result := 'SELECT * FROM COMMUN WHERE CO_TYPE="GCE"'
          + ' ORDER BY CO_CODE';
end;

function GetSqlGDPTT : string;
begin
  {$IFDEF PLUGIN} //compiler sans utilpgi
  Result := 'SELECT * FROM COMMUN WHERE CO_TYPE="GDP" ORDER BY CO_CODE';
  {$ELSE}
  Result := 'SELECT ' + GetSelectAll('CO') + ' FROM COMMUN WHERE CO_TYPE="GDP" ORDER BY CO_CODE';
  {$ENDIF !PLUGIN}
end;

function GetSqlGCXTT : string;
begin
  Result := 'SELECT * FROM COMMUN WHERE CO_TYPE="GCX" ORDER BY CO_CODE';
end;
{***********A.G.L.***********************************************
Auteur  ...... : C. BOUET
Créé le ...... : 09/05/2006
Modifié le ... : 09/05/2006
Description .. : le test seria est effectué après connexion à la base.
                 En cwas, c'est touours le cas, mais pas en 2 tiers.
                 Il faut donc désactiver ce test si l'on est pas connecté.
                 Le test de séria est maintenant toujours effectué dans le dispatch 16 avec :
                 - test seria de langue
                 - si échec (non sérialié ou limite du nb utilisateur atteinte,
                 	test séria en langue francaise
Mots clefs ... :
*****************************************************************}
{$IFDEF GCGC}
procedure laVariableGC.BeforeProtecAFF(sender : TObject);
begin
  // pas non encore connecté à la base, on désactive le test de séria
  V_PGI.NoProtec := not V_PGI.OkOuvert;
  if V_PGI.NoProtec then
    V_PGI.VersionDemo := True;
end;
{$ENDIF}

function GetCleDocByTob(TobXXX: Tob; const PrefixeSiPasTobReelle: String = ''): R_CleDoc;
var
  Prefixe: String;
begin
  with TobXXX do
  begin
    Prefixe := TableToPrefixe(NomTable);
    if Prefixe = '' then
      Prefixe := PrefixeSiPasTobReelle;
    { en cas d'exception, merci de faire : if Prefixe = 'ABC' then... }
    Result.NaturePiece := GetString (Prefixe + '_NATUREPIECEG');
    Result.Souche      := GetString (Prefixe + '_SOUCHE');
    Result.NumeroPiece := GetInteger(Prefixe + '_NUMERO');
    Result.Indice      := GetInteger(Prefixe + '_INDICEG');
    if TobXXX.FieldExists(Prefixe + '_NUMORDRE') then
      Result.NumOrdre    := GetInteger(Prefixe + '_NUMORDRE')
    else
      Result.NumOrdre    := 0;
  end
end;

{ Permet de savoir s'il l'on utilise la configuration par formulaires }
function wGereConfigurator: Boolean;
begin
  Result := GetParamSocSecur('SO_WEXTFORMS', false)
end;

// CCMX-CEGID ORGANIGRAMME PAIE POUR DEMANDES D'ACHATS
Function ExisteSalarieModulePaie :Boolean;
var
  Q: Tquery;
  LeNbre: Integer;
  st: string;
begin
  Result := False;
  St := 'SELECT COUNT(PSA_SALARIE) FROM SALARIES WHERE PSA_CIVILITE <> ""';
  Q := OpenSQL(st, TRUE,1,'', True);
  try
    if not Q.EOF then
    Begin
      LeNbre := Q.Fields[0].AsInteger;
      Result := LeNbre <> 0;
    end;
  finally
    ferme(Q);
  end;
end;

{$IFDEF GCGC} // BDU : Remplace GESCOM par GCGC
{$IFNDEF PAIEGRH}
{$IFNDEF PGIMAJVER}
{$IFNDEF CRM}
{$IFNDEF CMPGIS35}
procedure InitOrganigrammePaie;
Begin

  // POUR DEPORTSAL  VH_Paie.PGRESPONSABLES := True;
  VH_Paie.PGIncSalarie := True;
  VH_Paie.PgTypeNumSal := 'NUM';
  VH_Paie.PgLienRessource := True;
  VH_Paie.PGGESTIONCARRIERE := True;
  VH_Paie.PgSeriaFormation := TRUE;
  VH_Paie.PgeAbsences := TRUE;
end;
{$ENDIF}
{$ENDIF}
{$ENDIF !PGIMAJVER}
{$ENDIF !PAIEGRH}
// FIN CCMX-CEGID ORGANIGRAMME DA

{$IFDEF EAGLSERVER}
function VH_GC: LaVariableGC;
begin
  Result := LaVariableGC(RegisterVHSession('VH_GC', LaVariableGC))
end;
{$ENDIF EAGLSERVER}
{$ENDIF GCGC}

function GetInfoArtEcart (Article : String) : TOB;
var QArt : TQuery;
begin
  Result := nil;
  if Article = '' then Exit;
  if VH_GC.TOBArtEcart.GetString('GA_CODEARTICLE')<> Article then
  begin
		VH_GC.TOBArtEcart.InitValeurs(false);
		Qart := opensql('Select GA_CODEARTICLE,GA_ARTICLE, GA_LIBELLE from ARTICLE Where GA_CODEARTICLE="' + Article + '"', true,-1, '', True);
		if not Qart.eof then
    begin
			VH_GC.TOBArtEcart.SelectDB('',Qart);
      Result := VH_GC.TOBArtEcart;
    end;
    ferme (QArt);
  end else Result := VH_GC.TOBArtEcart;
end;

function GetInfoParSouche ( SoucheG,Champ : String ) : Variant ;
var StSQL : String;
    QQ    : TQuery;
begin
  result:='' ;
  if Champ = 'BS0_NUMMOISPIECE' then
  begin
    StSQL := 'SELECT ' + champ + ' FROM BSOUCHE WHERE BS0_TYPE="GES" AND BS0_SOUCHE="' + SoucheG + '"';
  end else
  begin
  StSQL := 'SELECT ' + champ + ' FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '"';
  end;
  QQ    := OpenSQL(StSQl, False);
  if not QQ.eof then
  begin
    if Champ = 'SH_LIBELLE' then
      result := TraduireMemoire(QQ.Findfield(Champ).AsString)
    else
      result := QQ.Findfield(Champ).AsVariant;
  end;

end;

end.

