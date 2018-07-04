unit CPTypeCons;

interface

uses
  HTB97, HEnt1;

const
{$IFNDEF CHR}
  MaxEche = 12 ;
{$ELSE}
  MaxEche = 50 ;
{$ENDIF}
  MaxAxe  = 5 ;
  MaxCatBud = 10 ;
  MaxSousPlan = 6 ;
  MaxTableLibre = 9 ;
  NBDECIMALTAUX = 6; {JP 25/05/05 : Nombre de décimale sur les taux de changes des devises}

  JalPurge : String = '---' ;
  CharJal : String = 'S' ;

  CodeISOFR = 'FR' ; //XMG 14/07/03
  CodeISOES = 'ES' ; //XMG 14/07/03
  CodeISOBE = 'BE' ; //SDA le 17/12/2007 version belge


type
  TExoDate = record
    Code              : string[3];
    Deb, Fin          : TDateTime;
    DateButoir,
    DateButoirRub,
    DateButoirBud,
    DateButoirBudgete : TDateTime;
    NombrePeriode     : Integer;
    EtatCpta          : string;
  end;

  TTabExo = array[1..5] of TExoDate;

  UneTaxe = record
    Regime               : string[3];
    Code                 : string[3];
    TauxACH, TauxVTE     : Double;
    CpteACH, CpteVTE     : string[17];
    EncaisACH, EncaisVTE : string[17];
  end ;

  LesTaxe = array[1..100] of UneTaxe ;

  TCPRoleCompta = (rcCollaborateur, rcReviseur, rcSuperviseur, rcExpert);

  TCPRevision = record
    Plan                 : string;
    DossierPretSupervise : Boolean;
    DossierSupervise     : Boolean;
    {$IFNDEF EAGLSERVER}
    BStatutRevision      : TToolBarButton97;
    {$ENDIF}
  end;

  // GCO - 08/07/2008
  TInfoBalSit = record
    CodeBal   : HString;
    Exo       : HString;
    Deb       : TDateTime;
    Fin       : TDateTime;
    TypeEcr   : HString;
    Supervise : Boolean;
    Historise : Boolean;
  end;

  TCPGSI = record
    OkGere   : Boolean;
    Journal  : HString;
    BalSitE  : TInfoBalSit;
    BalSitP  : TInfoBalSit;
    BalSitN2 : TInfoBalSit;
  end;
  // FIN GCO

  TSectRetri = (srOk, srNonStruct, srPasEnchainement, srBadArg);

  TStructure = (tstAux);

  TTableLibreCompta = (tlGene, tlAux, tlSect, tlCptBud, tlSectBud, tlEcrGen, tlEcrAna, tlEcrBud, tlImmo);

  TParamTableLibre = array[tlGene..tlImmo, 1..10] of Boolean;

  TStructureUniq = record
    CodeSp    : string;
    QuelCarac : string;
    Debu,Long : Integer;
  end;

  TInfoCpta = record
    Lg           : Byte;
    Cb           : Char;
    Chantier     : Boolean;
    Structure    : Boolean;
    Attente      : string; // LG 25/06/2007 - correction d'une fuite memmoire ds ChargeVHImmo
    AxGenAttente : string; // LG 25/06/2007 - correction d'une fuite memmoire ds ChargeVHImmo
    UneStruc     : TStructureUniq;
  end;

  {ne pas changer l'ordre des 5 premiers (axes)}
  TFichierBase = (fbAxe1,fbAxe2,fbAxe3,fbAxe4,fbAxe5,fbGene,fbAux,fbJal,fbBudGen,fbBudJal,
                 fbImmo,fbSect,fbCorresp,fbNatCpt,
                 fbBudSec1,fbBudSec2,fbBudSec3,fbBudSec4,fbBudSec5,
                 fbAxe1SP1,fbAxe1SP2,fbAxe1SP3,fbAxe1SP4,fbAxe1SP5,fbAxe1SP6,
                 fbAxe2SP1,fbAxe2SP2,fbAxe2SP3,fbAxe2SP4,fbAxe2SP5,fbAxe2SP6,
                 fbAxe3SP1,fbAxe3SP2,fbAxe3SP3,fbAxe3SP4,fbAxe3SP5,fbAxe3SP6,
                 fbAxe4SP1,fbAxe4SP2,fbAxe4SP3,fbAxe4SP4,fbAxe4SP5,fbAxe4SP6,
                 fbAxe5SP1,fbAxe5SP2,fbAxe5SP3,fbAxe5SP4,fbAxe5SP5,fbAxe5SP6,
                 fbRubrique,fbNone
                 ) ;

  ttTypePiece  = (tpReel, tpSim, tpPrev, tpSitu, tpRevi, tpCloture, tpIfrs);

  SetttTypePiece = set of ttTypePiece;

  TsetFichierBase = set of TFichierBase;

  TZoomTable = (tzGeneral,tzGCollectif,tzGNonCollectif,TzGCollClient,tzGCollFourn,
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
               tzCSPCorrespSec, //fb 26/06/2007 pour ANAL2
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

  TfourCpt = record
    Deb : string[17];
    Fin : string[17];
  end;

  TFBil = array[1..5] of TFourCpt;

  TFCha = array[1..5] of TFourCpt;

  TFPro = array[1..5] of TFourCpt;

  TFExt = array[1..2] of TFourCpt;

  TFF   = array[1..10] of TFourCpt;

  TLGTABLELIBRE  = array[1..MaxTableLibre, 1..10] of Integer;

  TNIVTABLELIBRE = array[1..MaxTableLibre, 1..10] of string[5];

  TSousPlan = record
    Code, Lib       : string;
    Debut, Longueur : Integer;
    ListeSP         : HTStringList;
  end;

  TSousPlanAxe = array[fbAxe1..fbAxe5, 1..MaxSousPlan] of TSousPlan;

  TUneCatBud = record
    fb        : TFichierBase;
    Code, Lib : string;
    SurJal    : array[1..MaxSousPlan] of string;
    SurSect   : array[1..MaxSousPlan] of string;
  end;

  TCatBud = array[1..MaxCatBud] of TUneCatBud;

  TSousPlanCat = array[1..MaxSousPlan] of TSousPlan ;

  TMPPop = Record
    MPGenPop, MPAuxPop, MPJalPop, MPExoPop : string;
    MPDatePop : tDAteTime ;
    MPNumPop, MPNumLPop, MPNumEPop : Integer ;
  end;

  tSauveSA_SaisieEff= array[0..15] of ShortInt;

  tCCMP = record
    LotCli,
    LotFou : Boolean;
  end;

  T_Eche = record
    ModePaie         : String3 ;
    DateEche         : TDateTime ;
    MontantP         : Double ; // Pivot
    MontantD         : Double ; // Devise
    Pourc            : Double ;
    ReadOnly         : Boolean ;
    Couverture       : Double ;
    CouvertureDev    : Double ;
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
    {$IFDEF CHR}
    Folio            :integer;
    Mtencais         : Double;
    Datecreation     : TDatetime;
    Utilisateur      : string3;
    {$ENDIF}

  end ;

  T_TabEche = array[1..MaxEche] of T_Eche ;

  T_ModeRegl = record
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
    EcartJours   : string ;
    ModifTva     : boolean ;
    TabEche      : T_TabEche ;
    {$IFDEF CHR}
    TabFolioP    : array[1..5] of Double ;
    TabFolioD    : array[1..5] of Double ;
    TabFolioE    : array[1..5] of Double ;
    {$ENDIF}
  end;

  T_Relance = class
    TypeRel,
    Famille : String3;
    Groupe, NonEchu,
    Scoring : boolean;
    Delais  : array[1..7] of Integer;
    Modeles : array[1..7] of String3;
  end;


  tQuelProduit = (execcs3, exeCCS5, execcs7, exeCCIMPEXP, exeCCMP, exeCCAUTO,
                  exePGIMAJVER, ExeAutre, ExeCTS3, ExeCTS5, exeCIIMPEXP, exeCCSTD, exeCOMSX) ;


  TTypeCalc=(Un, Deux, DeuxSansCumul, UnSurTableLibre, DeuxSurTableLibre, CPCpteJal,
             CPUn, CPUnUDF, UnBud, DeuxBud, UnNatEcr, DeuxNatEcr);



  ZSParams = record
    fb1,
    fb2        : TFichierBase ;
    SetTyp     : SetttTypePiece ;
    Multiple   : TTypeCalc ;
    SurQuelCpI,
    SurQuelCpO,
    Exo,
    Etab,
    Devise,
    TypeCumul,
    TypePlan   : string ;
    Date1,
    Date2      : TDateTime ;
    DevPivot,
    bSQLCumul  : boolean;
  end ;

  PZSParams = ^ZSParams ;

  TabDC = record
    TotDebit  : Double ;
    TotCredit : Double ;
  end;

  TabTot = array[0..7] of TabDC ;

// ==> Transféré depuis Saisbor.pas
type RParFolio = record
     ParPeriode,             // '01/03/1999'
     ParCodeJal,             // 'ODE'
     ParNumFolio        : string ;  // '77'
     ParNumLigne        : Integer ; // 3
     ParRecupLettrage   : boolean ; // si true on ouvre la saisie de bordereau pour les ecritures de types L
     ParGuide           : string ;
     ParCreatCentral    : boolean ; // appel en creation depuis le journal centralisateur
     ParCentral         : boolean ; // appel en modif depuis le journal centralisateur
     ParEta             : string ;
     end ;

implementation

end.




