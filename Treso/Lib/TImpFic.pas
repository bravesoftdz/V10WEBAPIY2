unit TimpFic ;

interface

Uses
    Classes,
    Hent1,
{$IFNDEF EAGL}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF EAGLSERVER}
    {$IFDEF EAGL}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    {$ENDIF EAGL}
{$ENDIF}
    uTOB
    ,ENT1
  , CPTypeCons
    ;

Const MaxCptLu=5000 ;
      PasDeBlanc : Boolean = FALSE ;
      Cblanc = '.' ;
      RecupSISCO : Boolean = FALSE ;
      RecupSISCOExt : Boolean = FALSE ;
      RecupSERVANT : Boolean = FALSE ;
      ImportMvtSU : Boolean = FALSE ;
      RecupPCL : Boolean = FALSE ;
      ImportBor : Boolean = TRUE ;

Const COTYPEIMP = 'IC1' ;

Type tModeJal = (mPiece,mBor,mLib) ;

type
  FMvtImport = ^TFMvtImport ;
  TFMvtImport = RECORD
//String
  IE_AFFAIRE,IE_ETATLETTRAGE,IE_LETTRAGE,IE_LETTREPOINTLCR,IE_LIBELLE,IE_REFEXTERNE,
  IE_REFINTERNE,IE_REFLIBRE,IE_REFPOINTAGE,IE_REFRELEVE,IE_RIB,IE_SECTION,
  IE_NUMPIECEINTERNE,IE_LIBRETEXTE0,IE_LIBRETEXTE1,IE_LIBRETEXTE2,IE_LIBRETEXTE3,
  IE_LIBRETEXTE4,IE_LIBRETEXTE5,IE_LIBRETEXTE6,IE_LIBRETEXTE7,IE_LIBRETEXTE8,
  IE_LIBRETEXTE9,IE_NUMEROIMMO : String ;
  IE_AXE : String[2] ;
// Combos
  IE_JOURNAL,IE_ENCAISSEMENT,IE_ECRANOUVEAU,IE_ETABLISSEMENT,IE_FLAGECR,IE_MODEPAIE,IE_NATUREPIECE,
  IE_REGIMETVA,IE_QUALIFPIECE,IE_QUALIFQTE1,IE_QUALIFQTE2,IE_SOCIETE,IE_TPF,IE_TVA,
  IE_TYPEANOUVEAU,IE_TYPEECR,IE_TYPEMVT,IE_DEVISE,IE_TABLE0,IE_TABLE1,
  IE_TABLE2,IE_TABLE3,IE_CONSO,IE_CODEACCEPT : String3 ;
// String17
  IE_AUXILIAIRE,IE_GENERAL,IE_CONTREPARTIEAUX,IE_CONTREPARTIEGEN : String17 ;
// entiers
  IE_CHRONO : Integer ;
  IE_NUMECHE,IE_NUMLIGNE,IE_NUMPIECE,
  IE_NUMVENTIL : Integer ;
// Dates
  IE_DATECOMPTABLE,IE_DATEECHEANCE,IE_DATEPAQUETMAX,IE_DATEPAQUETMIN,
  IE_DATEPOINTAGE,IE_DATEREFEXTERNE,IE_DATERELANCE,IE_DATETAUXDEV,
  IE_DATEVALEUR,IE_ORIGINEPAIEMENT,IE_LIBREDATE,IE_DATECREATION : TDateTime ;
// booleens
  IE_ECHE,IE_CONTROLE,IE_LETTRAGEDEV,IE_OKCONTROLE,
  IE_SELECTED,IE_TVAENCAISSEMENT,IE_TYPEANALYTIQUE,IE_VALIDE,IE_ANA,IE_INTEGRE,
  IE_LIBREBOOL0,IE_LIBREBOOL1
  (*,IE_LETTRAGEEURO,IE_SAISIEEURO*)  : String1 ;
// doubles
  IE_POURCENTAGE,IE_POURCENTQTE1,IE_POURCENTQTE2,IE_QTE1,IE_QTE2,
  IE_QUOTITE : Double ;
  IE_RELIQUATTVAENC,IE_TAUXDEV,IE_TOTALTVAENC,IE_COTATION : Double ;
  IE_DEBIT,IE_CREDIT,IE_CREDITDEV,
//  IE_CREDITEURO,
  IE_COUVERTURE,
  IE_COUVERTUREDEV,IE_DEBITDEV,
//  IE_DEBITEURO,IE_COUVERTUREEURO,
  IE_LIBREMONTANT0,IE_LIBREMONTANT1,IE_LIBREMONTANT2,IE_LIBREMONTANT3 : Double ;

  IE_ELEMENTARECUPERER : Boolean ;
  END ;

Type  TFmtEntete = RECORD
        FORMATDATE : String ;
        ANNEE4  : Boolean ;
        IMPORTATION  : Boolean ;
        SEPDATE  : Char ;
        ASCII : Boolean ;
        CRLF : Boolean ;
        MOINSDEVANT : Boolean ;
        TailleRec : Integer ;
        SEPDECIMAL : Char ;
        SEPMILLIER : Char ;
        DELIMITEUR : Char ;
        IGNORELIGNE : Integer ;
        FORMATSENS  : Integer ;
        NbLigne,NbChamps,Longueur : Integer ;
        PREFIXE,JOINTURES : String ;
        END ;

Type  TFmtDetail = RECORD
        NUMLIGNE : Integer ;
        CHAMP : String[40] ;
        DEBUT : Integer ;
        DROITE : Boolean ;
        LONGUEUR : Integer ;
        DECIMAL : Integer ;
        CORRESP : String ;
        END ;

Type TTabFmtDetail = Array[1..200] of TFmtDetail ;

Const
  _NEB_Comptable  = 'SOLDE COMPTABLE' ;
  _NEB_SoldeEnCoursNonEchu  = 'SOLDE EN COURS TRAITE NON ECHU' ;
  _NEB_SoldeEnCoursEchuNonRegle  = 'SOLDE EN COURS TRAITE ECHU NON REGLE' ;
  _NEB_DebitEnCoursNonEchu  = 'DEBIT EN COURS TRAITE NON ECHU' ;
  _NEB_DebitEnCoursEchuNonRegle  = 'DEBIT EN COURS TRAITE ECHU NON REGLE' ;
  _NEB_CreditEnCoursNonEchu  = 'CREDIT EN COURS TRAITE NON ECHU' ;
  _NEB_CreditEnCoursEchuNonRegle  = 'CREDIT EN COURS TRAITE ECHU NON REGLE' ;
  _NEB_Dev = 'DEVISE' ;

Type tLibreChampExp = (_EB_Comptable,_EB_SoldeEnCoursNonEchu,_EB_SoldeEnCoursEchuNonRegle,
                       _EB_DebitEnCoursNonEchu,_EB_DebitEnCoursEchuNonRegle,_EB_CreditEnCoursNonEchu,
                       _EB_CreditEnCoursEchuNonRegle) ;

Type tLibreCodeExp =(lCode,lLib) ;

Const LibLibreExp : Array[_EB_Comptable.._EB_CreditEnCoursEchuNonRegle,lCode..lLib] Of String =
  (('SOLDE COMPTABLE','Solde comptable'),
   ('SOLDE TRAITE NE','Solde en cours traite non échu'),
   ('SOLDE TRAITE ENR','Solde en cours traite échu non réglé'),
   ('DEBIT TRAITE NE','Débit en cours traite non échu'),
   ('DEBIT TRAITE ENR','Débit en cours traite échu non réglé'),
   ('CREDIT TRAITE NE','Crédit en cours traite non échu'),
   ('CREDIT TRAITE ENR','Crédit en cours traite échu non réglé')) ;

Type tMonnaieEcc = (EccPivot,EccOppose,EccDevise) ;

Type tMethodeCalcul = (Traite,TraiteEtFacture,Global) ;

Type tCritExpECC = Record
                   Aux1,Aux2 : String ;
                   DateCpta1,DateCpta2 : TDateTime ;
                   DateModif1,DateModif2 : TDateTime ;
                   OkPL,OkTL,OkTousMP : Boolean ;
                   DateButoir : TDateTime ;
                   Monnaie : tMonnaieEcc ;
                   NomFic : String ;
                   NomFicIni : String ;
                   St413EAR,St413EEP,St411,St411Solde : String ;
                   TL : Array[0..9] Of String ;
                   MethodeECTNE,MethodeECTENR : tMethodeCalcul ;
                   Format : String ;
                   DateNeg1,DateNeg2 : tDateTime ;
                   Masque : String ;
                   End ;

Type tTypeEnCours=(eccComptable,eccPortefeuille,eccPortefeuilleNonEchu,
                   eccPortefeuilleEchuNonregle,eccPortefeuilleEchuRegle,
                   eccChequeEnAttente) ;

Type tChampSup = Array[0..9] Of String ;

Type TDCExp = Record
              D,C,DE,CE : Variant ;
              End ;

Type tTabDCExp =Array[eccComptable..eccPortefeuilleEchuNonregle] Of tDCExp ;

Type tEnCours = Record
                Cpt,Cpt1 : Variant ;
                Montant : tTabDCExp ;
                Devise : Variant ;
                End ;

(*
Type tLibre = Record
              Champ : String ;
              Value : Variant ;
              End ;

Type tEnCours = Array[1..200] Of tLibre ;
*)
Type tFicEnCours = Record
                   NomFic : String ;
                   Nature,Format : String ;
                   entete : tFmtEntete ;
                   Detail : tTabFmtDetail ;
                   LR : HtStringList ; // LR : Resultats de la requète
                   End ;


Type tEnCoursP = Class
                 T : tEnCours ;
                 End ;

Type tResultImportCpte=(resRien,resCreer,resModifier) ;

Type TStSISCO = Record
                Jal,Date,TP,General,TC,AuxSect,Reference,Libelle,MP,Echeance,S,
                Montant,TE,NumP,Dev,TauxDev,CodeMontant,Montant2,Montant3,Etab,Axe,NumE,RefExterne,RefLibre,Qte,Affaire,DateRefExterne,QualQ2,RefReleve,RefP,
                // ajout me 16-12-2002
                lettre,datepaquetmin : String ;
                End ;

Type TFourchetteImport = Record
                         Cha,Pro,DebD,DebC,Imm,Ban,Cai,Cli,Fou,Sal,Dvs : TFCha ;
                         End ;
Type TCptCollSISCO = Record
                     Cpt,Aux1,Aux2 : String ;
                     Nat : String ;
                     End ;
Type TTabCptCollSISCO = Array[0..50] Of TCptCollSISCO ;
Type TCharRemplace = Record St1 : String ; St2 : String ; End ;

Type tNatImpExp =(nPiece,nRappro,nVirement,nPrelevement,nLcr,nTreso) ;
type TQFiche    = Array[0..5] Of TQuery ;
type TQDoublon  = Array[0..1] Of TQuery ;

Type TCptLu = RECORD
              Ok : Boolean ;
              Cpt : String[17] ;
              Collectif : String[17] ;
              CollTiers : String[17] ;
              Ventilable,VentilableDansFichier : Boolean ;
              Nature : String[3] ;
              RegimeTva : String[3] ;
              Pointable,Lettrable,OkMajSolde,EstColl : Boolean ;
              SoucheN,SoucheS,QualifPiece,ModeSaisie : String[3] ;
              Axe : String[5] ;
              LastDate : TDateTime ;
              LastNum,LastNumL : Integer ;
              DebitMvt,CreditMvt : Double ;
              TotDN,TotCN,TotDN1,TotCN1,TotDAno,TotCAno : Double ;
              Tva,TPF,TVAEnc,TvaEncHT : String[3] ;
              TP : String[17] ;
              IsTP : Boolean ;
              END ;

Type TCptLuP = Class
               T : TCptLu ;
               End ;

Type TTabCptLu = Array[1..MaxCptLu] Of TCptLu ;

Type PtTTabCptLu = ^TTabCptLu ;

Type tMasqueFic = Array[0..9] Of String ;

Type tTypeMasque = (mMasqueEcr,mMasqueAux) ;

Type tEnvImpAuto = Record
                   Rep,RepExport : String ;
                   VideRep : Boolean ;
                   MasqueFic : Array[mMasqueEcr..mMasqueAux] Of tMasqueFic ;
                   Import,Format,CodeFormat,User : String ;
                   EXE,ExeExport : String ;
                   Editeur : String ;
                   End ;


Type TScenario = Record
                 // pour Import
                 ImpAuto : Boolean ;
                 EstCharge : Boolean ;
                 Nature : String ;
                 RGen,RCli,RFou,RSal,RDiv,RSect1,RSect2,RSect3,RSect4,RSect5,
                 RCollCliT,RCollFouT,RMP,RMR,RGT : String ;
                 UseCorresp : Boolean ;
                 CorrespGen,CorrespAux,CorrespSect1,CorrespSect2,CorrespSect3,
                 CorrespSect4,CorrespSect5 : Boolean ;
                 TraiteCtr,TraiteTva : Boolean ;
                 ForceRIB,ANDetail : Boolean ;
                 Chemin,Prefixe,Suffixe : String ;
                 Doublon,ForcePiece : Boolean ;
                 DetruitFic : Boolean ;
                 CorrespMP,ValideEcr,ForceRibFou : Boolean ;
                 CorrespJal,CalcTauxDevOut : boolean ;
                 // Pour Export
                 FiltreGen,FiltreAux : Boolean ;
                 NatGen,NatAux : String ;
                 FourchetteImport : TFourchetteImport ;
                 ShuntPbAna : Boolean ;
                 Majuscule : Boolean ;
                 End ;

Type tOkEnrSISCO = Array[0..10] Of Boolean ;


Type TInfoImport = Record
                   ImportAuto : Boolean ;
                   Sc : TScenario ;
                   ImportOk : Boolean ;
                   NomFicRapportGlobal : String ;
                   NomFic : String ;
                   NomFicRejet : String ;
                   NomFicDoublon : String ;
                   NomFicRapport : String ;
                   NbLigIntegre : LongInt ;
                   NbLigLettre  : LongInt ;
                   NomFicOrigine : String ; // En cas de converion de fichier intermédiaire
                   NbPiece : Integer ;
                   TotDeb,TotCred : Double ;
                   LGenLu     : HTStringList ; { Liste de TCptLu }
                   LAuxLu     : HTStringList ;
                   LAnaLu     : HTStringList ;
                   LJalLu     : HTStringList ;
                   LMP        : HTStringList ;
                   LMR        : HTStringList ;
                   LRGT       : HTStringList ;
                   ListeCptFaux : TList ;
                   ListePieceFausse : TList ;
                   ListeEntetePieceFausse : HTStringList ;
                   ListePieceIntegre : HTStringList ; // Pour les pièces avec les N° après intégration (Tva Enc et contrepartie)
                   ListeEnteteDoublon : HTStringList ; // Pour les N° de pièce en doublons
                   CRListeEnteteDoublon : TList ; // Pour Les message d'erreurs des N° de pièce en doublons
                   ListeMvtTrouve : HTStringList ; // Liste des mouvements trouvés dans la base réceptrice : utilisé en lettrage uniquement
                   ListeDev : TList ; // Liste devise avec décimale et atux à prendre en compte en import
                   ListePbAna : HTStringList ; // Liste Des flux de IMPECR pour lesquel un problème d'équilibrage en analytique a été détecté
                   NbGenFaux,NbAuxFaux,NbAnaFaux,NbJalFaux : Integer ;
                   ForceNumPiece,ForcePositif,ForceBourrage,CtrlDb : Boolean ;
                   ForceQualif,Format,Lequel,CodeFormat,Prefixe,Table,FormatOrigine : String ;
                   DebutTraitement,FinTraitement : tDateTime ;
                   CodeSoc,NomSoc,TypeFic : String ; // Pour import auto avec gestion multi-société et format d'import déduit du fichier lui-même
                   EE : TEnvImpAuto ; // Pour import auto multi société
                   PbAna : Boolean ; // Si analytique absente sur un axe en vue du rééquilibrage post import
                   NomFicMvt,NomFicCpt : String ; // Pour Recup PCL
                   NomFicCptCGE,NomFicMvtCGE : String ; // Pour Recup PCL
                   OkEnr : TOkEnrSISCO ; // Pour Recup PCL
                   AuMoinsUnClient,AuMoinsUnFournisseur : Boolean ; // Pour récup PCL
                   OkFouCli,OkFouFou : Boolean ; // Pour récup PCL
                   PbEnr : tOkEnrSISCO ; // Pour récup PCL
                   ListePbEcrPCL : HTStringList ; // Pour récup PCL
                   RacineCptGenLet : String ; // Pour Récup PCL
                   RacineCptGenPoint : String ; // Pour Récup PCL
                   RacineJalAN : String ; // Pour récup PCL
                   ProfilPCL : String ; // Pour récup PCL
                   JalFaux : String ; // Pour récup cégid
                   OnFicheBase : Boolean ; // Pour Recup Cégid
                   CptECCSISCO : String ; // Pour récup cégid et PCL
                   BaseOrigine : String ; // Pour Recup Cégid
                   CorrigeAna : Boolean ;
                   LSoucheBOR : HtStringList ;
                   AuMoinsUnBordereau : Boolean ; // Pour import bordereau
                   AuMoinsUnePiece : Boolean ; // Pour import bordereau
                   ForceDevise : String ;
                   LettrageSISCOEnImport : Boolean ; // Pour récup SISCO PGE Uniquement
                   RacineJalVen : String ; // Pour récup SISCO PGE Uniquement
                   RacineJalAch : String ; // Pour récup SISCO PGE Uniquement
                   LCS : HtStringList ; // Code Stat sisco à créer. Pour récup SISCO PGE Uniquement
                   LGSISCO : Integer ; // Lg Compte SISCO. Pour récup SISCO PGE Uniquement
                   PointageAFaire : Boolean ; // pointage récupéré sur TRT. Pour récup SISCO PGE Uniquement
                   RecupODA : Boolean ; // Recup OD Analytique . Pour récup SISCO PGE Uniquement
                   AuMoinsUneODA : Boolean ; // Recup OD Analytique . Pour récup SISCO PGE Uniquement
                   StEEXBQ : String ; // Période avec des flux pointés(tokenisé) . Pour récup SISCO PGE Uniquement
                   TOBStat : TOB;   // Liste des codes statistiques sisco (enreg. 09)
                   NumFic : integer; // Numéro de fichier en cours de traitement
                   End ;
Type PtTInfoImport = ^TInfoImport ;

Type TIdentLigne = Record
     DP,CP,DD,CD,DE,CE,Q1,Q2,Debit,Credit : Double ;
     NumLig,NumEche,NumV : Integer ;
     NextNumVCEGID : Integer ; // Pour récup Cegid
     Gen,Aux,Sect,QualQ1,QualQ2,Axe,Sens : String ;
     Eche,Ana,ANouveau,ODAnal : Boolean ;
     END ;

type TIdentPiece = RECORD
     JalP,NatP,QualP : String3 ;
     DateP           : TDateTime ;
     NumP,Chrono     : Integer ;
     NumPDef         : Integer ;
     Doublon         : Boolean ;
     CodeMontant     : String ;
     TauxDev         : Double ;
     DecimDev        : Integer ;
     DernChronoE     : Integer ;
     DernChronoENonVentilable : Integer ;
     DernChronoEcrAna : Integer ;
     TotDP,TotCP,TotDD,TotCD,TotDE,TotCE : Double ;
     TotDPAna,TotCPAna,TotDDAna,TotCDAna,TotDEAna,TotCEAna : Double ;
     LigneEnCours,LignePrec : TIdentLigne ;
     FolioSISCO      : String ;
     NoFolioSISCO    : Integer ;
     SensDernLigneGen : Byte ; // 0 : Init  1 : Débit  2 : Crédit
     NumChronoEclate : Integer ; // Recup PCL
     END ;

Type TContact = Record
                Nom,Service,Fonction,Tel,Fax,Telex,RVA,Civilite,Principal : String ;
                End ;

Type TCptImport = Record
                  Cpt : String ;
                  Libelle : String ;
                  Nature : String ;
                  Lettrable : String ;
                  Collectif : String ;
                  EAN : String ;
                  T0,T1,T2,T3,T4,T5,T6,T7,T8,T9 : String ;
                  Adr1,Adr2,Adr3,CodePost,Ville : String ;
                  Rib : String ;
                  Axe : String ;
                  Pointable,Vent1,Vent2,Vent3,Vent4,Vent5,ModeSaisie : String ;
                  Pays,Abrege,Langue,MultiDevise,Devise,Tel,Fax,RegimeTVA,ModeRegl,Commentaire,Sens,Rep,TiersTP,IsTP,AvoirRBT,CptOrigine : String ;
                  Nif,Siret,APE,FormeJur : String ;
                  RibP : String ;
                  TiersMP,TiersEche : String ;
                  Deb,Lg : Integer ;
                  SoucheN,SoucheS : String ;
                  CptAuto,CptInt : String ; 
                  End ;

Type tDebFormat = Record
                  FormatFic : Integer ;
                  DebNum,DebMontant,DebSens,DebTC,DebNTP,DebGen : Integer ;
                  LgNum,LgGen : Integer ;
                  End ;

implementation

end.
