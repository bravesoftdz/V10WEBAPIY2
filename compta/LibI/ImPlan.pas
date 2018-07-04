unit ImPlan;

// ATTENTION : En cas de modifications dans les méthodes Charge, ChargePlan
//             ChargeMethode, Calcul et Sauve, effectuer la même modif.
//             dans les méthodes qui travaillent avec l'objet TOB - CA le 15/03/99
//             Ces méthodes se trouvent à la fin de ImPlan.pas
// 02/07/99 - CA - Modif GetVNCCedee
// CA - 02/07/99 - Si dernier jour du mois, on calcul sur le mois complet
// CA - 07/07/99 - Suppression AsDateTime sur GetValue
// CA - 08/07/1999 - Calcul de la base debut exo cession = base de la partie cédée.
// CA - 08/07/1999 - VNC Cedee par rapport à BAseDebExoCess
// CA - 19/07/1999 - Récupération du plan à une date donnée
// CA - 03/09/1999 - Minimum linéaire
// CA - 20/09/1999 - Ajout de la valeur d'achat ; recalcul de la reprise cédée.
// CA - 18/08/2000 - Remplacement DateFinAmort par DateMiseEnService
// CA - 18/08/2000 - Calcul date début amort suivant méthode.
// CA - 21/08/2000 - Calcul valeur résiduelle.
// CA - 14/05/2002 - Gestion du surmaortissement
// CA - 03/06/2003 - FQ 12109 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
// CA - 18/07/2003 - Suppression des champs euro
// CA - 20/01/2004 - FQ 13217 - Calcul de la réintégration après sortie
// FQ 15110 - CA  - 17/12/2004 - ExecuteSQL remplacé par ExisteSQL pour corriger problème dotation on calculée }
// FQ 15480 - MBo - 19/07/2005 - Non prise en cpte exceptionnel ds calcul de la VNC si reste a amortir apres date fin
//                              correction liée à FQ 15477 CA
// FQ 15630 - MBo - 29/07/2005 - si immo linéaire avec date d'achat = dernier jour du mois --> cumul antérieur faux
//                              (identique si date achat = 1er jour mois)
// FQ 16711 - MBo - 20/09/2005 - en cas de cession sur immo avec dotation exceptionnelle, l'exceptionnel
//                               était proratisé à tort
// -------- - Mbo - 25/10/2005 - modif de la proc getexcepExercice : ajout parametre bavecdeprec
//                                + calcul sur durée restante en mode éco uniquement
// FQ 16400 - MBO - 25/11/2005 - exceptionnel sur sortie non pris en compte dans le calcul de l'exceptionnel
// FQ 17102 - MBO - 29/11/2005 - l'exceptionnel doit impacter le plan fiscal au même titre que l'éco
// FQ 17154 - MBO - 06/12/2005 - pour les plan VAR : l'exceptionnel ne doit pas impacter le fiscal
// FQ 17215 - TGA - 21/12/2005 - GetParamSoc => GetParamSocSecur
// FQ 17229 - MBO - 23/12/2005 - pb de curseur non fermé dans cumul_depreciation
// -------- - MBO -   /02/2006 - impact antérieur dépréciation - nvelle f° calculRepriseDepreciation
// FQ 17547 - BTY -    03/2006 - Pb édition en Web Access à cause de VHIMMO
// FQ 15234 - MBO - 23/03/2006 - calcul du dégressif avec exercice < 12 mois qui n'est pas celui d'acquisition
//                               il manquait une annuité - modif CalculDateFinAmortissement
// -------- - MBO - 23/03/2006 - Alimentation de la variable datedebAmort dans calcul du linéaire qui
//                               n'était pas toujours initialisée (avertissement à la compil)
// -------- - TGA - 07/04/2006 Initialisation de champ à la création du plan
// -------- - MVG - 13/04/2006 - Correction avertissements à la compilation
// FQ 17569 - MBO - 27/04/2006 - Modif dates de debut amort - datedebeco et datedebfis
// FQ 18333 - MBO - 08/06/2006 - correction pb dans tableauDate si datedebfis < datedebeco
// FQ 18339 - MBO - 13/06/2006 - correction calcul du plan variable si datedebeco > N+1
// FQ 15234 - MBO - 15/06/2006 - Suite : on prend en compte l'annuité supplémentaire qu'à partir
//                               de l'exo qui suit celui qui est < 12 mois (sauf si exo debut amort eco)
// FQ 18471 - MVG - 22/06/2006 - Message d'erreur : Le variant ne référence pas un objet automation
// FQ 18544 - MBO - 11/07/2006 - exceptionnel pris 2 fois si date date fin amort théorique < date fin amort réélle
// FQ 15274 - MBO - 13/07/2006 - amortisst variable avec antérieurs
// FQ 18595 - MBO - 13/07/2006 - amortisst variable suite
// TGA 21/07/2006 Ajout codedpi et codedpiec
//          - BTY - 07/06      - Ajout DPI et DPIAffectee
//          - MBO - 24/08/2006 - correction conseil en compilation
//          - MVG - 08/06      - DPIAffectee, elle est non déductible
// FQ 18720   MVG - 31/08/2006 - le calcul se fait par rapport à la base éoonomique ou la base fisca mais pas sur la valeur d'achat
//          - MBO - 09/06      - Gestion des primes d'équipement
//          - MBO - 09/06      - Gestion des subventions d'équipement
//          - MBO - 25/10/2006 - Ajout d'une propriété dans les méthodes : PlanRef
//                               = 0 pour éco, fiscal et pri
//                               = 1 pour sbv si pas de plan fiscal ou NAM
//                               = 2 pour sbv si plan fiscal
//          - MBO - 11/2006    - Grosses modif dans la fonction RecupereSuivantDate (pour que les éditions à des dates
//                               < à certaines opérations soient justes
// FQ 19276 - MBO - 06/12/2006 - pb impression SBV si cession et impression avant date cession
// FQ 19530 - MBO - 16/12/2006 - si dotation au delà de fin d'amortissement : la dotation était réinitialisée
// chantiel fiscal : 2 nouvelles valeur pour plan ref :
//                   3 = pour dérogatoire
//                   4 = pour réintégration
// FQ 20256 BTY - 05/2007 En Traitements/Recalcul des plans, enregistrer le type de dérogatoire
// BTY 06/2007 - Cumul_depreciation rendue publique
// mbo 29/06/07 - modif de la fonction getminimumlineaire (ajout de la duree en entrée) suite g° fiscale
// mbo 10/09/07 - fq 21405 - cession sur un exercice postérieur à l'exercice de fin d'amortissement
// BTY Attention Cedele = i_datecession et pas i_datecreation, et reconduire l'info sur la mùéthode Copie
// BTY 09/2007 FQ 21346 fPlanUO publié pour optimiser l'édition des amortissements variables
// MBO 05/10/2007 - ajout dans la classe TPlanAmort du paramètre Qte (quantite) - fq 19172
// MBO 29/10/2007 - fq 21754 - ajout paramètre à TraiteSortie pour prise en cpte ou non du param 'amort du jour de cession'
interface

uses
  SysUtils,
  Controls,
  HEnt1, HCtrls, Math, UTob, ImPlanMeth, AmType, (*Ent1,*) ParamSoc
{$IFDEF eAGLClient}
{$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
{$ENDIF eAGLClient}
{$IFDEF SERIE1}
  , S1Util  //YCP 25/08/05
{$ELSE}
  , Ent1
{$ENDIF}
 ;

const
  MethodeEco: integer = 1;
  MethodeFisc: integer = 2;

  MethodePri: integer = 3;
  MethodeSbv: integer = 4;

  // ajout pour chantier fiscal
  MethodeDerog : integer = 5;
  MethodeReint : integer = 6;

type
  TPlanAmort = class
    CodeImmo: string;
    LibelleImmo: string;
    AmortEco,AmortFisc: TMethodeAmort;

    AmortPri: TMethodeAmort;
    AmortSbv: TMethodeAmort;
    //chantier fiscal
    AmortDerog, AmortReint : TMethodeAmort;

    CompteImmo: string;
    DateMiseEnService: TDateTime;
    DateAchat: TDateTime;
    Fiscal: boolean;
    TableauDate: array[0..MAX_LIGNEDOT] of TDateTime;
    TabTypeExc: array[0..MAX_LIGNEDOT] of string;
    ValReintegration: double;
    TauxQuotePart: double;
    DateCreation: TDateTime;
    DateCalcul: TDateTime;
    NumSeq: integer;
    NumSeqPrec: integer;
    TypeOpe: string;
    Modifie: boolean;
    VNC: double;
    RevisionFisc: double;
    RevisionEco: double;
    MontantExc: double;
    TypeExc: string;
    BaseDebutExoEco: double;
    BaseDebutExoFisc: double;
    NbMoisReprise: integer;
    BaseDebutAmort: double;
    BaseFinAmort: double;
    CoeffReintegration: double;
    DecalageAffDotations: integer;
    ValeurHT: double;
    ValeurAchat: double;
    TypeAmortCession: string;
    DateAmortCession: TDateTime;
    ImmoCloture: boolean;
    SortiePlusExcep: boolean;
    ChangementPlan: boolean;
    bDotationExercice: boolean;
    ValeurTVARecuperable,
    ValeurTvaRecuperee: double;
    Suramort: boolean;
    fbRecalculReprise : boolean;

    CumulDepre : Double;    //cumul depre
    RepriseDepre : Double;  // reprise maximale possible recalculé et transmis par la cloture lors du changt exo
    DureeRest : boolean;    // calcul sur durée restante suite saisie depreciation
    GestDeprec : boolean;   // true si dépréciation au cours exo précédents
    VncRestEco : boolean;      // calcul du plan économique futur avec la vnc
    journala  : string;     // = *** si calcul du plan futur avec la vnc
    EcoVar : boolean;         // fq 17154 - ajout mbo 6.12.05
    CodeDPI : String;
    CodeDPIEC : String;
// BTY 03/06 PB en édition Web Access cause VHIMMO => définition publique pour ImouPlan qui s'en sert
{$IFDEF EAGLSERVER}
    fEncours: TExoDate;
{$ENDIF}

     // fq 17596 - mbo - nouvelles dates début amortissement
     DateDebEco : TdateTime;
     DateDebFis : TdateTime;
     // BTY 07/06 Zones DPI
     DPI : boolean;
     DPIAffectee : double;

     // mbo - 01.09.06 - gestion des primes d'équipement et des subventions
     PRI : boolean;
     MNTPrime : double;
     PRITaux :double;
     PRIDuree:integer;
     PRIMethode:string;
     PRIDateDeb : TdateTime;

     SBV : boolean;
     MNTSbv : double;
     SBVTaux : double;
     SBVDuree : integer;
     SBVMethode : string;
     SBVDateDeb : TDateTime;
     SBVDateSBV : TDateTime;
     RepriseCedeeSBV : double;

     GestionFiscale : boolean;  // ajout pour chantier fiscal
     VncRestFisc : boolean;   // calcul du fiscal sur la vnc fiscale
     RepriseDerog : double;
     RepriseReint : double;
     RepriseDRCedee : double;
     RepriseRICedee : double;

     EstRemplacee : boolean;   // ajout suite opération de remplacement
     Remplace : boolean;

     CedeLe : TDateTime;    // FQ 21405 - ajout mbo 10.09.07 = I_DATECESSION
     Qte : double;     // FQ 19172 - ajout mbo 05.10.07 = I_QUANTITE

  private
    { Déclarations privées }
    fPlanUO: TOB; { Plan d'unités d'oeuvre }
    fTotalUO: double; { Nombre total d'unités d'oeuvre }
// BTY 03/06
//{$IFDEF EAGLSERVER}
//    fEncours: TExoDate;
//{$ENDIF}

   // ajout mbo 15.06.06 FQ 15234
   // Utilisé dans le calcul du dégressif - est positionné à true si
   // un exercice d'amortissement est < inférieur à 12 mois et que cet exercice n'est pas
   // l'exercice de début d'amortissement
    AjoutPeriode : boolean;


    procedure InitPlan;
    procedure InitMethodes;
    procedure ChargePlan(Q: TQuery);
    procedure ChargeMethode(Q: TQuery);
    procedure ChargePlanTOB(OB: TOB);
    procedure ChargeMethodeTOB(OB: TOB);
    procedure Affecte(QPlan: TQuery; TPlan: TOB = nil);
    procedure NettoiePlan;
    procedure CalculRepriseVariable(MethodeAmort: TMethodeAmort;
      DateDebut: TDateTime);
    // BTY 02/06 Publiques pour ImedCalc anciennes éditions ouvertes à CWAS
    //function DateFinEnCours: TDateTime;
    //function DateDebutEnCours: TDateTime;


    // BTY 07/06 DPI
    function MontantDPIAffectee (CodeImmo : string) : double;
  public
    { Déclarations publiques }
    constructor Create(Initialisation: boolean);
    destructor destroy; override;
    procedure Init;
    procedure Charge(Q: TQuery);
    procedure ChargeTOB(OB: TOB);
    procedure Copie(Source: TPlanAmort);
    function Sauve: boolean;
    procedure SauveTOB(OBPlan: TOB);
    procedure Delete;
    function Open(tCodeImmo, tNumPlan: string; TAmort: TOB = nil): TQuery;
    function Recupere(tCodeImmo, tNumPlan: string): boolean;
    procedure Calcul(Q: TQuery; tDateCalcul: TDateTime);
    procedure CalculTOB(OB: TOB; tDateCalcul: TDateTime; bRecalculReprise : boolean = False);
    procedure CalculPlanCede(T: TOB; tDateCalcul: TDateTime);  // non utlisée 
    procedure Recalcul(Q: TQuery);
    procedure RecalculTOB(OB: TOB);
    procedure CalculReprise(MethodeAmort: TMethodeAmort);
    procedure CalculDotation(MethodeAmort: TMethodeAmort; tDateCalcul:
      TDateTime; bCloture: boolean);
    procedure CalculLineaire(MethodeAmort: TMethodeAmort; tDateCalcul:
      TDateTime; bCloture: boolean);
    procedure CalculDegressif(MethodeAmort: TMethodeAmort; tDateCalcul:
      TDateTime; bCloture: boolean);
    procedure CalculVariable(MethodeAmort: TMethodeAmort; DateCalcul:
      TDateTime; bCloture : boolean);
    procedure CalculDateFinAmortissement(MethodeAmort: TMethodeAmort);
    procedure CalculDateFinAmortReprise(MethodeAmort: TMethodeAmort);
    procedure SetBaseAmortFin(MontantOpe: double);
    procedure SetReprises(PlanOrig: TPlanAmort; Valeur: double);
    procedure SetTypeOpe(NewTypeOpe: string);
    procedure SetTypeExc(NewTypeExc: string);
    function GetBaseFinAmort: double;
    function GetDateFinAmortEx(MethAmort: TMethodeAmort): TDateTime;
    function ProrataUOReste (DateCalcul:TdateTime): double;  // ajout mbo 7.11.05
    function GetVNCAvecMethode(MethAmort: TMethodeAmort; DateRef: TDateTime):double;
    function GetRestePri(MethAmort: TMethodeAmort; DateRef: TDateTime):double;
    function GetResteSbv(MethAmort: TMethodeAmort; DateRef: TDateTime):double;

    function GetVNCCEdee(MethAmort: TMethodeAmort; DateRef: TDateTime): double;
    procedure BasculeDotationCedee(PCede: TPlanAmort; DateOpe: TDateTime);

    procedure GetCumulsDotExercice(DateRef: TDateTime; var CumulEco, CumulFisc:
      double; bCession, bReprise, bClot: boolean);

    procedure GetCumulsDotExerciceDR(DateRef: TDateTime; var CumulDR, CumulFEC:
      double; bCession, bReprise, bClot: boolean); // FQ 17512 chantier fiscal


    procedure GetCumulPRI(DateRef: TDateTime; var CumulPri:double; bCession, bReprise, bClot: boolean);
    procedure GetCumulSBV(DateRef: TDateTime; var CumulSbv:double; bCession, bReprise, bClot: boolean);

    function GetDotationEncours(Methode: TMethodeAmort; bCession: boolean):
      double;
    function GetDotationExercice(DateFinExe: TDateTime; Methode: TMethodeAmort;
      bCession: boolean): double;
    function GetCessionExercice(tDateCalcul: TDateTime; Methode: TMethodeAmort):
      double;
    procedure SetRevisionEco(Montant: double);
    procedure SetRevisionFisc(Montant: double);
    procedure SetMontantExc(Valeur: double);
    procedure RestaureRepriseSurImmoEclate(PEclate: TPlanAmort);
    procedure CalculRepriseLineaire(Methode: TMethodeAmort; DateDebutPer,
      tDateCalcul: TDateTime);
    procedure CalculRepriseDegressif(Methode: TMethodeAmort; tDateCalcul:
      TDateTime);
    function GetDateFinAmort: TDateTime;
    procedure CalculReprises;
    function GetDotationGlobale(TypeMethode: integer; tDateCalcul: TDateTime; var
      dDot, dCes, dExc: double): double;
    // mbo 24.10.05 function GetExcepExercice(tDateCalcul: TDateTime; var DateRevisionNam:
    //  TDateTime; bAvecSortie: boolean = false): double;
    function GetExcepExercice(tDateCalcul: TDateTime; var DateRevisionNam:
      TDateTime; bAvecSortie: boolean = false; bAvecDeprec:boolean = false): double;
    procedure GetElementExceptionnel(var Dot, DotExc: double; var sTypeExc:
      string);
    function ReintegrationFiscale: boolean;
    function TraiteMontantDotationProrata(DateOp, tDateCalcul: TDateTime;
      MontantDot: double): double;
    procedure AffecteDotationCedee(PCedee: TPlanAMort; tDateCalcul: TDateTime);
    procedure CalculProrataDotationCedee(MethodeAmort: TMethodeAmort;
      tDateCalcul: TDateTime);
    procedure CalculExeptionnel(MethodeAmort: TMethodeAmort; tDateCalcul:
      TDateTime);
    procedure SetInfosCession(TypeAmort: string; DateCession: TDateTime);
    procedure GetCumulDotationCession(var CumulcesEco, CumulCesFis: double;
      tDateREf: TDateTime);

    procedure GetCumulDotCessionPRI(var CumulCesPri: double;tDateREf: TDateTime);
    procedure GetCumulDotCessionSBV(var CumulCesSBV: double;tDateREf: TDateTime);

    procedure ResetTableauDot(var Methode: TMethodeAmort; DateDebut: TDateTime);
    function GetBaseExercice(DateFinExe: TDateTime; Methode: TMethodeAmort):
      double;
    function GetBaseCessionExercice(DateFinExe: TDateTime; Methode:
      TMethodeAmort): double;
    procedure BasculeBaseCession(MethodeEco: TMethodeAmort);
    procedure RecupereSuivantDate(tDateCalcul: TDateTime; TAmort: TOB = nil;
      TLog: TOB = nil);
    procedure CalculExceptionnelSurSortie(MethodeAmort: TMethodeAmort;
      tDateCalcul: TDateTime);
    procedure CalculProrataDotationOrigine(MethodeAmort: TMethodeAmort; Prorata:
      double);
    procedure CalculPourTraitement(TypeTraitement: string; CLoCumulDepre: double);
    function GetDateDebutAmort(MethodeAmort: TMethodeAmort): TDateTime;
    function GetExcepAnter(tDebExo: TDateTime): double;
    procedure UpdateBasePlan;
    function ExoAchatInferieurA12(DateDeb: TDateTime): boolean;  // mbo 27.04.026 fq 17569
    function ExoInferieurA12(DateDebEx:TdateTime): boolean;  // mbo 23.03.06 fq 15234
    procedure ChargePlanUO(T: TOB);
    // impact antérieur dépréciation mbo 02.06
    procedure CalculRepriseTheorique(MethodeAmort: TMethodeAmort; tDateCalcul: tDateTime);

    function CalculRepriseDepreciation(MethodeAmort: TMethodeAmort;
                    tDateDeb:TDateTime; tDateFin:TDateTime; Clot:boolean):double;
    function DateFinEnCours: TDateTime;     // BTY 02/06
    function DateDebutEnCours: TDateTime;   // BTY 02/06
    function Recup_Baseeco(Codeimmo:String;bsortie:boolean): double;
    function Recup_Basefisc(Codeimmo:String;bsortie:boolean): double;
    function Recup_RepriseCedeeSBV(Codeimmo:String;bsortie:boolean): double;  //MBO 03/11/06

    //chantier fiscal
    function Determine_gestion_fiscale(Q: TQuery; RepriseEco, RepriseFisc:double):boolean;
    procedure Calcul_Derog_Reint;
    Procedure CumulDerogReint(tDateCalcul:TdateTime;var CumDerog, CumReint: double);
    procedure CalculReprise_Derog_Reint;
    function GetValResiduelle(CodeImmo:string; MethAmort: TMethodeAmort; DateRef: TDateTime; bCession:boolean): double;

    // mbo 10.05.07 cumul brut des dépréciations saisies pour export des immos
    function CumulDepreciationBrut(Codeimmo:String; MethodeAmort: TMethodeAmort): double;

{$IFDEF EAGLSERVER}
    procedure SetDatesServer(Encours: TExoDate);
{$ENDIF}
    // modif mbo 29/10/2007 - fq 21754 - ajout paramètre amortJsortie (NON si pas amort jour de cession)
    procedure TraiteSortie(DateOperation: TDateTime; ModeCalcul : string; MntExcSortie : double; stTypeExc : string;
                           AmortJsortie : string = '');

    function CalculDureeReprise(PlanUtil: string) : integer;
    function DegressifTauxReprise(MethodeAmort: TMethodeAmort; tDateCalcul:
      TDateTime; bCloture: boolean):double;

    // BTY fonction à rendre publique
    // mbo 27.10.05 procedure Cumul_depreciation(Codeimmo:String;Cumul:Double);
    function Cumul_depreciation(Codeimmo:String; bCloture:boolean): double;

    published
       property TotalUO : double read fTotalUO;
       property PlanUO : TOB read fPlanUO;   // FQ 21346
  end;

  function RecalculSauvePlan(Q: TQuery): integer;
  function RecalculSauvePlanTOB(OB: TOB): integer;
  procedure RecuperePlanRepriseTheorique(Q: TQuery; var CumEco, CumFisc: double);
  function GetMinimumLineaire(Q: TQuery; duree:integer): double;
  procedure UpdateBaseImmo(bConfirme: boolean = true; vbRecalculReprise : boolean = false);
  function CtrlImmoOpe: boolean;

implementation

uses ImOuPlan
      , HStatus
      , ImEnt
      , HMsgBox
      , Outils   // FQ 20256
      ;
// calcul la reprise d'amortissement
// PRINCIPE : pour une methode, calcul le montant deja amorti entre Datedebut et dateref

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 25/05/2004
Modifié le ... :   /  /
Description .. : Calcul de la reprise dans le cas d'un amortissement variable
Mots clefs ... :
*****************************************************************}

procedure TPlanAmort.CalculRepriseVariable(MethodeAmort: TMethodeAmort;
  DateDebut: TDateTime);
var
  i: integer;
  CumulUOAmorti : double;
begin
  i := 0;
  CumulUOAmorti := 0;
  MethodeAmort.Reprise := 0;
  while ((i < fPlanUO.Detail.count) and (fPlanUO.Detail[i].GetValue('IUO_DATE')
    < DateDebut)) do
  begin
    CumulUOAmorti := CumulUOAmorti +  fPlanUO.Detail[i].GetValue('IUO_UNITEOEUVRE');
    Inc(i);
  end;
  if (fTotalUO<>0) then
    MethodeAmort.Reprise := Arrondi(MinValue([MethodeAmort.Base * (CumulUOAmorti/fTotalUO), MethodeAmort.Base]), V_PGI.OkDecV)
  else MethodeAmort.Reprise := 0;
end;

procedure TPlanAmort.CalculRepriseLineaire(Methode: TMethodeAmort; DateDebutPer,
  tDateCalcul: TDateTime);

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/02/2004
Modifié le ... : 16/02/2004
Description .. : Calcul du nombre de jour entre 2 dates en considérant que
Suite ........ : les mois ont 30 jours (une année fait 360 jours)
Suite ........ : On considère DateDebut inclus et DateFin non inclus.
Mots clefs ... :
*****************************************************************}
  function NombreJour(DateDebut, DateFin: TDateTime): integer;
  var
    PremMois, PremAnnee, nMois: Word;
    nJourDebut, nJourFin: integer;
    dd, df, mm, yy: Word;
  begin
    { Calcul du nombre de jours des mois entiers }
    NombreMois(FinDeMois(DateDebut) + 1, DebutDeMois(DateFin) - 1, PremMois,
      PremAnnee, nMois);
    { Calcul du nombre de jours pour le mois de début }
    DecodeDate(DateDebut, yy, mm, dd);
    if (dd = DaysPerMonth(yy, mm)) then
      // mbo FQ 15630 - 29/07/2005 nJourDebut := 30
      nJourDebut := 1
    else
      nJourDebut := (30 - dd) + 1;
    { Calcul du nombre de jours pour le mois de fin }
    DecodeDate(DateFin, yy, mm, df);
    if (df = DaysPerMonth(yy, mm)) then
      nJourFin := 30
    else
      nJourFin := df - 1;
    { Calcul du nombre de jour total }
    if (FinDeMois(DateDebut) = FinDeMois(DateFin)) then
      // Début et fin de période sur le même mois
      Result := df - dd
    else
      Result := nJourDebut + (nMois * 30) + nJourFin;
  end;

var
  Prorata: double;
  PremMois, PremAnnee, NbMois: Word;
  YDeb, MDeb, DDeb, DFin, MFin, YFin: Word;
  NbMoisPeriode,
//  NbMoisAmort,
  NbJourDetention: integer;
begin
  Prorata := 0.00;
  NOMBREMOIS(DateDebutPer, Methode.DateFinAmort, PremMois, PremAnnee, NbMois);
  NbMoisPeriode := NbMois;
  NOMBREMOIS(DateDebutPer, tDateCalcul, PremMois, PremAnnee, NbMois);
//  NbMoisAmort := MaxIntValue([NbMois - 1, 0]);
  DecodeDate(DateDebutPer, YDeb, MDeb, DDeb);
  DecodeDate(tDateCalcul - 1, YFin, MFin, DFin);
  if (DFin = DaysPerMonth(YFin, MFin)) then
    DFin := 30;
      // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if (DDeb = DaysPerMonth(YDeb, MDeb)) then
    DDeb := 30;
      // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  //  NbJourDetention := (((NbMoisAmort-1)*30)+(DFin-MinIntValue([30,DDeb])+1));
  NbJourDetention := NombreJour(DateDebutPer, tDateCalcul);
  // CA - 16/02/2004 - Si NbJourDetention<0, on ne fait pas de calcul (pas de sens) : cas des dates de début <> 1er jour du mois
  if ((NbJourDetention > 0) and (NbMoisPeriode <> 0)) then
    Prorata := ((NbMoisPeriode * 30) / (Methode.Duree * 30)) * (NbJourDetention
      / (NbMoisPeriode * 30));
  Methode.Reprise := Arrondi(MinValue([Methode.Base * Prorata, Methode.Base]),
    V_PGI.OkDecV);
end;

// calcul la reprise d'amortissement
// PRINCIPE : pour une methode, calcul le montant deja amorti entre Datedebut et datecalcul

procedure TPlanAmort.CalculRepriseDegressif(Methode: TMethodeAmort; tDateCalcul:
  TDateTime);
var
  Cumul: double;
  i: integer;
begin
  bDotationExercice := True;
  CalculDotation(Methode, GetDateDebutAmort(Methode), False);
  i := 0;
  Cumul := 0.00;
  while (Methode.TableauDot[i] <> 0) and (TableauDate[i] <= tDateCalcul) do
  begin
    Cumul := Cumul + Methode.TableauDot[i];
    inc(i);
  end;
  Methode.Reprise := Arrondi(MinValue([Cumul, Methode.Base]), V_PGI.OkDecV);
end;

//Modification Reprise ECO ou FISCAL

function TPlanAmort.GetDateDebutAmort(MethodeAmort: TMethodeAmort): TDateTime;
begin
  { fq 17569 - mbo
  if MethodeAmort.Methode = 'DEG' then
    result := DateAchat
  else
    result := DateMiseEnService;
  }
  if methodeAmort.planUtil = 'FISC' then
     result := DateDebFis
  else if methodeAmort.planUtil = 'ECO' then
     result := DateDebEco
  else if methodeAmort.PlanUtil = 'PRI' then
     result := PRIDateDeb
  else if methodeAmort.PlanUtil = 'SBV' then
     result := SBVDateDeb
  else
     result := DateDebEco;  // pour chantier fiscal

end;

//constructeur de la classe TPlanamort

constructor TPlanAmort.Create(Initialisation: boolean);
begin
  inherited Create;
  AmortEco := TMethodeAmort.Create;
  AmortFisc := TMethodeAmort.Create;

  AmortPri := TMethodeAmort.Create;
  AmortSbv := TMethodeAmort.Create;

  AmortDerog := TMethodeAmort.Create;
  AmortReint := TMethodeAmort.Create;


  fPlanUO := TOB.Create('', nil, -1);
  if Initialisation then
    Init;
end;

destructor TPlanAmort.destroy;
begin
  fPlanUO.Free;
  AmortFisc.free;
  AmortEco.free;

  AmortPri.free;
  AmortSbv.free;

  inherited;
end;

procedure TPlanAmort.copie(Source: TPlanAmort );
var
  i: integer;
begin
  CodeImmo := Source.CodeImmo;
  LibelleImmo := Source.LibelleImmo;
  CompteImmo := Source.CompteImmo;
  Fiscal := Source.Fiscal;
  DateMiseEnService := Source.DateMiseEnService;
  DateAchat := Source.DateAchat;
  DateCreation := Source.DateCreation;
  DateCalcul := Source.DateCalcul;
  NumSeq := Source.NumSeq;
  NumSeqPrec := Source.NumSeqPrec;
  TypeOpe := Source.TypeOpe;
  Modifie := Source.Modifie;
  VNC := Source.VNC;
  MontantExc := Source.MontantExc;
  TypeExc := Source.TypeExc;
  BaseDebutExoEco := Source.BaseDebutExoEco;
  BaseDebutExoFisc := Source.BaseDebutExoFisc;
  NbMoisReprise := Source.NbMoisReprise;
  BaseDebutAmort := Source.BaseDebutAmort;
  BaseFinAmort := Source.BaseFinAmort;
  ValReintegration := Source.ValReintegration;
  CoeffReintegration := Source.CoeffReintegration;
  TauxQuotePart := Source.TauxQuotePart;
  DecalageAffDotations := Source.DecalageAffDotations;
  ValeurHT := Source.ValeurHT;
  ValeurTvaRecuperable := Source.ValeurTvaRecuperable;
  ValeurTvaRecuperee := Source.ValeurTvaRecuperee;
  ValeurAchat := Source.ValeurAchat;
  TypeAmortCession := Source.TypeAmortCession;
  DateAmortCession := Source.DateAmortCession;
  SortiePlusExcep := Source.SortiePlusExcep;
  ImmoCloture := Source.ImmoCloture;
  ChangementPlan := Source.ChangementPlan;
  Suramort := Source.Suramort;
  for i := low(TableauDate) to high(TableauDate) do
    TableauDate[i] := Source.TableauDate[i];
  AmortEco.Copie(TPlanAmort(Source).AmortEco, 'ECO');   // ajout mbo 25.10.05 planeco
  AmortFisc.Copie(TPlanAmort(Source).AmortFisc, 'FISC');  // ajout mbo 25.10.05 planfiscal

  AmortPri.Copie(TPlanAmort(Source).AmortPri, 'PRI');  // ajout mbo 01.09.06 planprime
  AmortSbv.Copie(TPlanAmort(Source).AmortSbv, 'SBV');  // ajout mbo 28.09.06 plan Subvention

  AmortDerog.Copie(TplanAmort(Source).AmortDerog, 'DRG'); // ajout mbo fq 17512
  AmortReint.Copie(TplanAmort(Source).AmortReint, 'REI'); // ajout mbo fq 17512

  // fq 17569 - mbo
  DateDebEco := Source.DateDebEco;
  DateDebFis := Source.DateDebFis;

  // BTY 07/06 DPI
  DPI := Source.DPI;
  DPIAffectee := Source.DPIAffectee;

  // mbo - 01.09.06 - prime

  PRI := Source.PRI;
  MNTPrime := Source.MNTPrime;
  PRITaux:=Source.PRITaux;
  PRIDuree:=Source.PRIDuree;
  PRIMethode:=Source.PRIMethode;
  PRIDateDeb:=Source.PRIDateDeb;

  SBV := Source.SBV;
  MNTSbv := Source.MNTSbv;
  SBVTaux:=Source.SBVTaux;
  SBVDuree:=Source.SBVDuree;
  SBVMethode:=Source.SBVMethode;
  SBVDateDeb:=Source.SBVDateDeb;
  SBVDateSBV:=Source.SBVDateSBV;

  // FQ 21405 date de cession
  CedeLe := Source.CedeLe;
  Qte := Source.Qte; //fq 19172 Quantite

end;

procedure TPlanAmort.RecalculTOB(OB: TOB);
var
  i: integer;
begin
  for i := 0 to MAX_LIGNEDOT do
    TableauDate[i] := iDate1900;
  InitMethodes;
  CalculTOB(OB, iDate1900);
end;

procedure TPlanAmort.Init;
begin
  InitPlan;
  InitMethodes;
end;

// Initialisation du plan - Valeurs par défaut

procedure TPlanAmort.InitPlan;
var
  i: integer;
begin
  NumSeq := 0;
  Modifie := false;
  TypeOpe := 'CRE';
  NbMoisReprise := 0;
  CoeffReintegration := 0.00;
  TypeAmortCession := '';
  DateAmortCession := iDate1900;
  SortiePlusExcep := false;
  ChangementPlan := false;
  DecalageAffDotations := 0;
  ImmoCloture := false;
  for i := 0 to MAX_LIGNEDOT do
    TableauDate[i] := iDate1900;
  for i := 0 to MAX_LIGNEDOT do
    TabTypeExc[i] := ' ';
end;

//Initialisation des méthode d'amortissement - Valeurs par défaut

procedure TPlanAmort.InitMethodes;
begin
  AmortEco.RazMethodeAmort;
  AmortFisc.RazMethodeAmort;

  AmortPri.RazMethodeAmort;
  AmortSbv.RazMethodeAmort;

  // ajout chantier fiscal
  AmortDerog.RazMethodeAmort;
  AmortReint.RazMethodeAmort;


end;

// Chargement du plan à partir d'une Query - Appel principal

procedure TPlanAmort.Charge(Q: TQuery);
begin
  ChargePlan(Q);
  ChargeMethode(Q);
end;

// Chargement du plan à partir d'une Query

procedure TPlanAmort.ChargePlan(Q: TQuery);
Var
  base : double;
  bfiscal : boolean;
begin
  NumSeq := Q.FindField('I_PLANACTIF').AsInteger;
  CodeImmo := Q.FindField('I_IMMO').AsString;
  LibelleImmo := Q.FindField('I_LIBELLE').AsString;
  CompteImmo := Q.FindField('I_COMPTEIMMO').AsString;
  NbMoisReprise := Q.FindField('I_DUREEREPRISE').AsInteger;
  DateAchat := StrToDate(Q.FindField('I_DATEPIECEA').AsString);
  DateMiseEnService := StrToDate(Q.FindField('I_DATEAMORT').AsString);
  BaseDebutExoEco := Q.FindField('I_BASEECO').AsFloat;
  BaseDebutExoFisc := Q.FindField('I_BASEFISC').AsFloat;
  BaseDebutAmort := Q.FindField('I_BASEAMORDEBEXO').AsFloat;
  BaseFinAmort := Q.FindField('I_BASEAMORFINEXO').AsFloat;
  DateCreation := Q.FindField('I_DATECREATION').AsDateTime;
  MontantExc := Q.FindField('I_MONTANTEXC').AsFloat;
  TypeExc := Q.FindField('I_TYPEEXC').AsString;
  ValReintegration := Q.FindField('I_REINTEGRATION').AsFloat;
  TauxQuotePart := Q.FindField('I_QUOTEPART').AsFloat / 100.0;
  ValeurHT := Q.FindField('I_MONTANTHT').AsFloat;
  ValeurTvaRecuperable := Q.FindField('I_TVARECUPERABLE').AsFloat;
  ValeurTvaRecuperee := Q.FindField('I_TVARECUPEREE').AsFloat;
  ValeurAchat := Q.FindField('I_VALEURACHAT').AsFloat;
  ImmoCloture := Q.FindField('I_ETAT').AsString <> 'OUV';

  // Ajout mbo 10.09.07 - FQ 21405
  CedeLe := Q.FindField('I_DATECESSION').AsDateTime;

  // Ajout mbo 05.10.07 - FQ 19172
  Qte := Q.FindField('I_QUANTITE').AsFloat;

   // BTY 07/06 DPI
  DPI := (Q.FindField('I_DPI').AsString = 'X');
  DPIAffectee := MontantDPIAffectee (CodeImmo);

  // MVG 30/08/2006 - FQ 18720 : le calcul se fait par rapport à la base éoonomique ou la base fiscale
  Base :=Recup_Baseeco(Codeimmo,Q.FindField('I_QUANTITE').AsFloat = 0.0);
  bFiscal := Q.FindField('I_METHODEFISC').AsString <> '';
  if bfiscal then Base :=Recup_Basefisc(Codeimmo,Q.FindField('I_QUANTITE').AsFloat = 0.0);

  //  if Q.FindField('I_MONTANTHT').AsFloat<> 0.00 then
  //    CoeffReintegration := (Q.FindField('I_REINTEGRATION').AsFloat/Q.FindField('I_MONTANTHT').AsFloat); //YCP Prendre TTC-TVA récupéré
    // CA - 20/01/2004 - FQ 13217 - Calcul de la réintégration après sortie

// MVG 30/08/2006 - FQ 18720 : le calcul se fait par rapport à la base éoonomique

  if Base <> 0.00 then
    CoeffReintegration := ((Q.FindField('I_REINTEGRATION').AsFloat + DPIAffectee) / Base);
      //YCP Prendre TTC-TVA récupéré
  ChangementPlan := Q.FindField('I_OPECHANGEPLAN').AsString = 'X';
  Suramort := Q.FindField('I_SURAMORT').AsString = 'X';

  { Chargement du plan d'unités d'oeuvre }
  fPlanUO.LoadDetailFromSQL('SELECT * FROM IMMOUO WHERE IUO_IMMO="' + CodeImmo +
    '"');
  fTotalUO := Q.FindField('I_UNITEOEUVRE').AsFloat;

  //Cumul_depreciation(Codeimmo,cumuldepre);
  RevisionEco := Q.FindField('I_REVISIONECO').AsFloat;

  // TGA Code DPI de l'immo
  CodeDPI   := Q.FindField('I_DPI').AsString;
  CodeDPIEC := Q.FindField('I_DPIEC').AsString;

  // fq 17476  - mbo 13.02.06 - impact dépréciation sur plan fiscal
  RevisionFisc := Q.FindField('I_REVISIONFISCALE').AsFloat;

  CumulDepre := Cumul_depreciation(Codeimmo,false);

  journala  := Q.FindField('I_JOURNALA').AsString ;
  IF Q.FindField('I_JOURNALA').AsString = '***' then
     VncRestEco := True
  Else
     VncRestEco := False;

  VncRestFisc := (Q.FindField('I_FUTURVNFISC').AsString = '***'); // chantier fiscal

  // impact antérieur dépréciation
  IF (Q.FindField('I_REVISIONECO').Asfloat <> 0) OR
      (Q.FindField('I_REPRISEDEP').Asfloat <> 0) then   // mbo 02.06
   DureeRest := True
  ELSE
   DureeRest := False;

  // ajout mbo fq 18339
  if DureeRest = true then
  begin
     if Q.FindField('I_DATEDEBECO').AsDateTime > VHImmo^.Encours.Fin then
     begin
        DureeRest := false;
        VncRestEco := true;
        VncRestFisc := true;  // ajout pour chantier fiscal
     end;
  end;
  // fin fq 18339

  // fq 17154 - ajout mbo 6.12.2005
  if Q.FindField('I_METHODEECO').AsString = 'VAR' then
     EcoVar := true
  else
     EcoVar := false;

  // fq 17569 - ajout mbo 27.04.2006
  DateDebEco := StrToDate(Q.FindField('I_DATEDEBECO').AsString);
  DateDebFis := StrToDate(Q.FindField('I_DATEDEBFIS').AsString);

  // mbo - 01.09.06 - prime d'équipement

  PRI := (Q.FindField('I_SBVPRI').AsFloat <> 0) or (Q.FindField('I_SBVPRIC').AsFloat <> 0);
  if Q.FindField('I_SBVPRI').AsFloat <> 0 then
     MNTPrime:=(Q.FindField('I_SBVPRI').AsFloat)
  else
     MNTPrime:=(Q.FindField('I_SBVPRIC').AsFloat);

  SBV := (Q.FindField('I_SBVMT').AsFloat <> 0) or (Q.FindField('I_SBVMTC').AsFloat <> 0);

  // modif mbo 25.10.06 if Q.FindField('I_SBVMT').AsFloat <> 0 then
  MNTSbv:=(Q.FindField('I_SBVMT').AsFloat);
  //else
  //  MNTSbv:=(Q.FindField('I_SBVMTC').AsFloat);

  if (PRI) then
  begin
     if (bFiscal) then
     begin
        PRIDuree:=Q.FindField('I_DUREEFISC').AsInteger;
        PRITaux:=Q.FindField('I_TAUXFISC').AsFloat;
        PRIMethode:=Q.FindField('I_METHODEFISC').AsString;
        PRIDateDeb := DateDebFis;
     end else
     begin
        if (Q.FindField('I_METHODEECO').AsString <> 'NAM') then
        begin
           PRIDuree:=Q.FindField('I_DUREEECO').AsInteger;
           PRITaux:=Q.FindField('I_TAUXECO').AsFloat;
           PRIMethode:=Q.FindField('I_METHODEECO').AsString;
           PRIDateDeb := DateDebEco;
        end else
        begin
           PRIDuree:=Q.FindField('I_CORVRCEDDE').AsInteger;
           PRITaux:= GetTaux('LIN',DateMiseEnService, DateMiseEnService, PRIDuree);
           PRIMethode:='LIN';
           PRIDateDeb := StrToDate(Q.FindField('I_DATEAMORT').AsString);
        end;
     end;
  end;

  if (SBV) then
  begin
     SBVDateSBV := StrToDate(Q.FindField('I_SBVDATE').AsString);
     RepriseCedeeSBV :=Recup_RepriseCedeeSBV(Codeimmo,Q.FindField('I_QUANTITE').AsFloat = 0.0);

     if (bFiscal) then
     begin
        SBVDuree:=Q.FindField('I_DUREEFISC').AsInteger;
        SBVTaux:=Q.FindField('I_TAUXFISC').AsFloat;
        SBVMethode:=Q.FindField('I_METHODEFISC').AsString;
        SBVDateDeb := DateDebFis;
     end else
     begin
        if (Q.FindField('I_METHODEECO').AsString <> 'NAM') then
        begin
           SBVDuree:=Q.FindField('I_DUREEECO').AsInteger;
           SBVTaux:=Q.FindField('I_TAUXECO').AsFloat;
           SBVMethode:=Q.FindField('I_METHODEECO').AsString;
           SBVDateDeb:= DateDebEco;
        end else
        begin
           SBVDuree:=Q.FindField('I_CORVRCEDDE').AsInteger;
           SBVTaux:= GetTaux('LIN',DateMiseEnService, DateMiseEnService, SBVDuree);
           SBVMethode:='LIN';
           SBVDateDeb:= StrToDate(Q.FindField('I_DATEAMORT').AsString);
        end;
     end;
  end;

  // ajout pour chantier fiscal

  GestionFiscale := (Q.FindField('I_NONDED').AsString = 'X');
  RepriseDerog := Q.FindField('I_REPRISEDR').AsFloat;
  RepriseReint := Q.FindField('I_REPRISEFEC').AsFloat;
  RepriseDRCedee := Q.FindField('I_REPRISEFDRCEDEE').AsFloat;
  RepriseRICedee := Q.FindField('I_REPRISEFECCEDEE').AsFloat;

  EstRemplacee := (Q.FindField('I_REMPLACEE').AsString <> '');
  Remplace := (Q.FindField('I_REMPLACE').AsString <> '');

end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 27/07/2006
Modifié le ... :   /  /
Description .. : Somme des enregistrements de la table IMMOMVTD pour
Suite ........ : l'immo indiquée en entrée
Mots clefs ... :
*****************************************************************}
function TPlanAmort.MontantDPIAffectee (CodeImmo : string) : double;
var Q: TQuery;
begin
  result := 0;

  if DPI then
     begin
       Q := OpenSQL('SELECT SUM(IZ_MONTANT) AS MTDPI FROM IMMOMVTD WHERE IZ_IMMO="' + CodeImmo + '"', True) ;
       try
         result := Q.FindField('MTDPI').AsFloat;
       finally
         ferme(Q);
       end;
     end;
end;


// Cumul des dépréciations déjà effectuées
// procedure TPlanAmort.Cumul_depreciation(Codeimmo:String, Cumul:Double);
{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 07/11/2005
Modifié le ... :   /  /
Description .. : renseigne le booleen GestDeprec = saisie de dépréciation
Suite ........ : sur exercices précédents
Suite ........ : renseigne CumulDepre suivant que l'on est en clôture
Suite ........ : d'exercice ou non
Mots clefs ... :
*****************************************************************}
function TPlanAmort.Cumul_depreciation(Codeimmo:String;bcloture:boolean): double;
Var QLog: TQuery;
    type_op : string;
    CdeSql : string;
    Cumul : Double;

begin
  GestDeprec := false;
  Cumul :=0.00;

  // 1ère partie : sert à renseigner le boolean Gestdeprec = saisie d'une dépréciation d'actif
  //               sur exercices précédents
  type_op:='DPR';

  if bCloture then
    CdeSql := 'SELECT IL_MONTANTDOT FROM IMMOLOG WHERE (IL_IMMO="'
        + CodeImmo + '" AND IL_TYPEOP ="' + type_op+'")'

    //QLog := OpenSQL('SELECT IL_MONTANTDOT FROM IMMOLOG WHERE (IL_IMMO="'
    //    + CodeImmo + '" AND IL_TYPEOP ="' + type_op+'")' , True)
  else
    CdeSql := 'SELECT IL_MONTANTDOT FROM IMMOLOG WHERE (IL_IMMO="'
        + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'" AND IL_DATEOP<"' +
        USDateTime(DateDebutEnCours) +'")';

    //     QLog := OpenSQL('SELECT IL_MONTANTDOT FROM IMMOLOG WHERE (IL_IMMO="'
    //    + CodeImmo + '" AND IL_TYPEMODIF ="' + type_op +'" AND IL_DATEOP<="' +
    //    USDateTime(DateFinEnCours) +'")' , True);

  {while not QLog.Eof do     // optimisation du code 23.12.2005
  begin
      GestDeprec := true;
      QLog.Next;
  end;
  Ferme(QLog);  // fq 17229  - mbo 23.12.2005
  }

  GestDeprec := ExisteSQL (CdeSql);


  // 2ème partie : sert à alimenter le cumuldepre
  //               à la cloture on prend directement RepriseDepre qui est calculé par immoclo
  //               et on positionne DureeRest à faux
  //               si on est pas en cloture d'exercice : on lit immolog pour récupérer le
  //               montant dans il_revisioneco
  if Bcloture then
  begin
     cumul := RepriseDepre;
     DureeRest := false;
  end else
    begin
    type_op:='CLO';

    QLog := OpenSQL('SELECT IL_REVISIONECO FROM IMMOLOG WHERE (IL_IMMO="'
        + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'" AND IL_DATEOP<"' +
        USDateTime(DateDebutEnCours) +'")' , True);
    try
      while not QLog.Eof do
      begin
        Cumul :=  QLog.FindField('IL_REVISIONECO').AsFloat;
        QLog.Next;
      end;
    finally
      Ferme(QLog);   // fq 17229
    end;
  end;

  Result := cumul;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 10/05/2007
Modifié le ... :   /  /    
Description .. : Cette fonction renvoie un cumul depreciation brut =
Suite ........ : antérieur dépréciation saisi +
Suite ........ : le montant des dépréciations saisis au cours de la vie de
Suite ........ : l'immo à fin d'exercice n-1
Mots clefs ... :
*****************************************************************}
function TPlanAmort.CumulDepreciationBrut(Codeimmo:String; MethodeAmort: TMethodeAmort): double;
Var QLog: TQuery;
    type_op : string;
    Cumul : Double;

begin
  Cumul := MethodeAmort.RepriseDep;

  type_op:='DPR';

  QLog := OpenSQL('SELECT IL_MONTANTDOT FROM IMMOLOG WHERE (IL_IMMO="'
        + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'" AND IL_DATEOP<"' +
        USDateTime(DateDebutEnCours) +'")' , True);
  try
    while not QLog.Eof do
    begin
      Cumul :=  Cumul + QLog.FindField('IL_MONTANTDOT').AsFloat;
      QLog.Next;
    end;
  finally
    Ferme(QLog);
  end;

  Result := cumul;
end;

//---------------------------------------------------------------------------------------------

// Récupération de la base d'amortissement fiscale sur une immo
// pour les immos sorties pour lesquelles I_BASEFISC est à 0
// MVG 30/08/2006 - FQ 18720

{***********A.G.L.***********************************************
Auteur  ...... : mvg
Créé le ...... : 30/08/2006
Modifié le ... :   /  /
Mots clefs ... :
*****************************************************************}
function TPlanAmort.Recup_Baseeco(Codeimmo:String;bsortie:boolean): double;
Var QLog: TQuery;
    type_op : string;
    Cumul : Double;
begin
  Cumul :=0.00;

  if bsortie then
  begin
    type_op:='CES';
      QLog := OpenSQL('SELECT IL_BASEECOAVMB FROM IMMOLOG WHERE (IL_IMMO="'
        + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'")', True);
    try
      while not QLog.Eof do
      begin
        Cumul :=  QLog.FindField('IL_BASEECOAVMB').AsFloat;
        QLog.Next;
      end;
    finally
      Ferme(QLog);
    end;
  end
  else
    begin
    cumul:=BaseDebutExoEco;
    end;
  Result := cumul;
end;

// Récupération de la base d'amortissement et fiscale sur une immo
// pour les immos sorties pour lesquelles I_BASEFISC est à 0
// MVG 30/08/2006 - FQ 18720
{***********A.G.L.***********************************************
Auteur  ...... : mvg
Créé le ...... : 30/08/2006
Modifié le ... :   /  /
Mots clefs ... :
*****************************************************************}
function TPlanAmort.Recup_Basefisc(Codeimmo:String;bsortie:boolean): double;
Var QLog: TQuery;
    type_op : string;
    Cumul : Double;
begin
  Cumul :=0.00;

  if bsortie then
  begin
    type_op:='CES';
      QLog := OpenSQL('SELECT IL_BASEFISCAVMB FROM IMMOLOG WHERE (IL_IMMO="'
        + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'")', True);
    try
      while not QLog.Eof do
      begin
        Cumul :=  QLog.FindField('IL_BASEFISCAVMB').AsFloat;
        QLog.Next;
      end;
    finally
      Ferme(QLog);
    end;
  end
  else
  begin
    cumul:=BaseDebutExoFisc;
  end;
  Result := cumul;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 03/11/2006
Modifié le ... :   /  /    
Description .. : Récupère dans immolog le montant de la reprise 
Suite ........ : SUBVENTION qui a été remis à 0 lors de la cession
Mots clefs ... :
*****************************************************************}
function TPlanAmort.Recup_RepriseCedeeSBV(Codeimmo:String;bsortie:boolean): double;
Var QLog: TQuery;
    type_op : string;
    Cumul : Double;
begin
  Cumul :=0.00;

  if bsortie then
  begin
    type_op:='CES';
      QLog := OpenSQL('SELECT IL_MONTANTAVMB FROM IMMOLOG WHERE (IL_IMMO="'
        + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'")', True);
    try
      while not QLog.Eof do
      begin
        Cumul :=  QLog.FindField('IL_MONTANTAVMB').AsFloat;
        QLog.Next;
      end;
    finally
      Ferme(QLog);
    end;
  end
  else
    begin
    cumul:=0.00;
    end;
  Result := cumul;
end;


// Chargement des méthodes d'amortissement à partir d'une Query

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 09/10/2006 - nouveautés de la version 7
Modifié le ... :   /  /
Description .. : les 2 derniers éléments de la méthode sont
Suite ........ : Ces 2 éléments servent pour les méthodes PRI et SBV
Suite ........ : Pour les méthodes ECO et FISC ils sont à false
Suite ........ :
Suite ........ : creation = est positionné à true lors de la saisie de la prime
Suite ........ :                 ou de la subvention
Suite ........ : ReprisePrem = ne sert que pour la subvention
Suite ........ :               si = true ---> la reprise est cumulée sur la première
Suite ........ :                             dotation calculée
Suite ........ :              si false ---> la première dotation est calculée
Suite ........ :                            normalement et les antérieurs saisis
Suite ........ :                            sont bien à part
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.ChargeMethode(Q: TQuery);
var
  Clot: boolean;
  ReprisePrem : boolean;
begin
  Clot := Q.FindField('I_ETAT').AsString <> 'OUV';

  // ajout mbo 09.10.06 pour reprise subvention sur la première dotation calculée
  if Q.FindField('I_DPIEC').AsString = '-' then
     ReprisePrem := false
  else
     ReprisePRem := true;

  AmortEco.AffecteMethodeAmort(TableauDate,
    Q.FindField('I_DATECESSION').AsDateTime,
    Q.FindField('I_REVISIONECO').AsFloat,
    0.00 (*Exceptionnel*), Q.FindField('I_REPRISEECO').AsFloat,
    Q.FindField('I_REPCEDECO').AsFloat, Q.FindField('I_BASEECO').AsFloat,
    Q.FindField('I_TAUXECO').AsFloat, Q.FindField('I_DUREEECO').AsInteger,
    Q.FindField('I_METHODEECO').AsString, Clot, 'ECO', (*mbo 25.10.05 ajout planeco*)
    Q.FindField('I_REPRISEDEP').AsFloat,(*Tga 27/01/2006*)
    false,  (*mbo 03.10.06 paramètre creation*)
    false, (*mbo 09.10.06 paramètre cumul anterieur sur 1ère dotation*)
    0);  // plan de référence

  Fiscal := Q.FindField('I_METHODEFISC').AsString <> '';

  // fq 15274 - on ne prend pas en compte la reprise dépréciation sur le fiscal
  // si la méthode éco = VAR
  if EcoVar = false then
     AmortFisc.AffecteMethodeAmort(TableauDate,
       Q.FindField('I_DATECESSION').AsDateTime,
       Q.FindField('I_REVISIONFISCALE').AsFloat,
       0.00 (*Exceptionnel*), Q.FindField('I_REPRISEFISCAL').AsFloat,
       Q.FindField('I_REPCEDFISC').AsFloat, Q.FindField('I_BASEFISC').AsFloat,
       Q.FindField('I_TAUXFISC').AsFloat, Q.FindField('I_DUREEFISC').AsInteger,
       Q.FindField('I_METHODEFISC').AsString, Clot, 'FISC', (*mbo 25.10.05 ajout planfisc*)
       Q.FindField('I_REPRISEDEP').AsFloat, false,(*Tga 27/01/2006 + mbo 03.10.06 *)
       false, 0) // ajout mbo 09.10.06 cumul anterieur sur première dotation
  else
     AmortFisc.AffecteMethodeAmort(TableauDate,
       Q.FindField('I_DATECESSION').AsDateTime,
       Q.FindField('I_REVISIONFISCALE').AsFloat,
       0.00 (*Exceptionnel*), Q.FindField('I_REPRISEFISCAL').AsFloat,
       Q.FindField('I_REPCEDFISC').AsFloat, Q.FindField('I_BASEFISC').AsFloat,
       Q.FindField('I_TAUXFISC').AsFloat, Q.FindField('I_DUREEFISC').AsInteger,
       Q.FindField('I_METHODEFISC').AsString, Clot, 'FISC',0, false,
       false, 0); // ajout pour cumul antérieur sur première dotation

   // ajout mbo - 01.09.06
   PRI := ((Q.FindField('I_SBVPRI').AsFloat <> 0.00) or (Q.FindField('I_SBVPRIC').AsFloat <> 0.00));
   if (fiscal) then
   begin
      AmortPri.AffecteMethodeAmort(TableauDate,
         Q.FindField('I_DATECESSION').AsDateTime,
         0.00 (*cumul dépréciation*),
         0.00 (*Exceptionnel*),
         Q.FindField('I_REPRISEUO').AsFloat,
         Q.FindField('I_REPRISEUOCEDEE').AsFloat,
         Arrondi((Q.FindField('I_SBVPRI').AsFloat/2),V_PGI.OkDecV),
         Q.FindField('I_TAUXFISC').AsFloat, Q.FindField('I_DUREEFISC').AsInteger,
         Q.FindField('I_METHODEFISC').AsString, Clot, 'PRI',0, false,
         false, 0);   // ajout pour cumul anterieur première dotation
   end else
   begin
      if (Q.FindField('I_METHODEECO').AsString = 'NAM') then
      begin
         AmortPri.AffecteMethodeAmort(TableauDate,
         Q.FindField('I_DATECESSION').AsDateTime,
         0.00 (*cumul dépréciation*),
         0.00 (*Exceptionnel*),
         Q.FindField('I_REPRISEUO').AsFloat,
         Arrondi((Q.FindField('I_REPRISEUOCEDEE').AsFloat),V_PGI.OkDecV),
         Arrondi((Q.FindField('I_SBVPRI').AsFloat/2),V_PGI.OkDecV),
         GetTaux('LIN', Q.FindField('I_DATEAMORT').AsDateTime, Q.FindField('I_DATEAMORT').AsDateTime,
                  Trunc(Q.FindField('I_CORVRCEDDE').AsFloat)),
         Trunc(Q.FindField('I_CORVRCEDDE').AsFloat),
         'LIN', Clot, 'PRI', 0.00, false, false, 0);
      end else
      begin
         AmortPri.AffecteMethodeAmort(TableauDate,
         Q.FindField('I_DATECESSION').AsDateTime,
         0.00 (*cumul dépréciation*),
         0.00 (*Exceptionnel*),
         Q.FindField('I_REPRISEUO').AsFloat,
         Arrondi((Q.FindField('I_REPRISEUOCEDEE').AsFloat),V_PGI.OkDecV),
         Arrondi((Q.FindField('I_SBVPRI').AsFloat/2),V_PGI.OkDecV),
         Q.FindField('I_TAUXECO').AsFloat, Q.FindField('I_DUREEECO').AsInteger,
         Q.FindField('I_METHODEECO').AsString, Clot, 'PRI',0, false, false, 0);  // on met reprisedep à 0
      end;
   end;

   SBV := (Q.FindField('I_SBVMT').AsFloat <> 0.00) or (Q.FindField('I_SBVMTC').AsFloat <> 0.00);

   if (fiscal) then
   begin

      AmortSbv.AffecteMethodeAmort(TableauDate,
         Q.FindField('I_DATECESSION').AsDateTime,
         0.00 (*cumul dépréciation*),
         0.00 (*Exceptionnel*),
         Q.FindField('I_CORRECTIONVR').AsFloat,
         RepriseCedeeSBV,
         Arrondi((Q.FindField('I_SBVMT').AsFloat),V_PGI.OkDecV),
         Q.FindField('I_TAUXFISC').AsFloat, Q.FindField('I_DUREEFISC').AsInteger,
         Q.FindField('I_METHODEFISC').AsString, Clot, 'SBV',0, false,
         ReprisePrem, 2);
   end else
   begin
      if (Q.FindField('I_METHODEECO').AsString = 'NAM') then
      begin
         AmortSbv.AffecteMethodeAmort(TableauDate,
         Q.FindField('I_DATECESSION').AsDateTime,
         0.00 (*cumul dépréciation*),
         0.00 (*Exceptionnel*),
         Q.FindField('I_CORRECTIONVR').AsFloat,
         RepriseCedeeSBV,
         Arrondi((Q.FindField('I_SBVMT').AsFloat),V_PGI.OkDecV),
         GetTaux('LIN', Q.FindField('I_DATEAMORT').AsDateTime, Q.FindField('I_DATEAMORT').AsDateTime,
                  Trunc(Q.FindField('I_CORVRCEDDE').AsFloat)),
         Trunc(Q.FindField('I_CORVRCEDDE').AsFloat),
         'LIN', Clot, 'SBV', 0.00, false,  (*reprise depreciation*)
         ReprisePrem, 0);
      end else
      begin
         AmortSbv.AffecteMethodeAmort(TableauDate,
         Q.FindField('I_DATECESSION').AsDateTime,
         0.00 (*cumul dépréciation*),
         0.00 (*Exceptionnel*),
         Q.FindField('I_CORRECTIONVR').AsFloat,
         RepriseCedeeSBV,
         Arrondi((Q.FindField('I_SBVMT').AsFloat),V_PGI.OkDecV),
         Q.FindField('I_TAUXECO').AsFloat, Q.FindField('I_DUREEECO').AsInteger,
         Q.FindField('I_METHODEECO').AsString, Clot, 'SBV',0, false,   // on met reprisedep à 0
         ReprisePrem, 1);
      end;
    end;

    // ajout pour chantier fiscal

    AmortReint.AffecteMethodeAmort(TableauDate,
    Q.FindField('I_DATECESSION').AsDateTime,
    0.00 (*dotation depreciation*),
    0.00 (*Exceptionnel*), Q.FindField('I_REPRISEFEC').AsFloat,
    Q.FindField('I_REPRISEFECCEDEE').AsFloat, 0.00 (*base*),
    0.00 (*taux*), 0 (*duree*),
    '' (*methode*), Clot, 'REI',
    0.00 (*reprise depreciation*),
    false,  (*mbo 03.10.06 paramètre creation*)
    false, (*mbo 09.10.06 paramètre cumul anterieur sur 1ère dotation*)
    0);  // plan de référence

    AmortDerog.AffecteMethodeAmort(TableauDate,
    Q.FindField('I_DATECESSION').AsDateTime,
    0.00 (*dotation depreciation*),
    0.00 (*Exceptionnel*), Q.FindField('I_REPRISEDR').AsFloat,
    Q.FindField('I_REPRISEFDRCEDEE').AsFloat, 0.00 (*base*),
    0.00 (*taux*), 0 (*duree*),
    '' (*methode*), Clot, 'DRG',
    0.00 (*reprise depreciation*),
    false,  (*mbo 03.10.06 paramètre creation*)
    false, (*mbo 09.10.06 paramètre cumul anterieur sur 1ère dotation*)
    0);  // plan de référence

end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 01/03/2004
Modifié le ... :   /  /
Description .. : Enregistrement du plan d'amortissement
Suite ........ : - 01/03/2004 : réécriture en TOB pour compatibilité eAGL
Mots clefs ... :
*****************************************************************}

function TPlanAmort.Sauve: boolean;
var
  DatePrec, DateFinAmort: TDateTime;
  Methode: TMethodeAmort;
  Lig: integer;
  CessEco, CessFisc, CumulEco, CumulFisc, DotDerog, DotEco, DotFisc: double;
  DotPri, CessPri: double;
  DotSbv, CessSbv: double;

  TAmort, T: TOB;
  D2, D3, D4 : double;
  DateDebut, DateFin : TDateTime;

  // ajout pour chantier fiscal
  DotReint : double;

begin
DatePrec := iDate1900;

DateDebut := iDate1900;
DateFin   := iDate1900;

TAmort := TOB.Create('', nil, -1);
try
  NumSeq := NumSeq + 1;
  Modifie := false;
  CumulEco := AmortEco.Reprise;
  CumulFisc := AmortFisc.Reprise;

  DateFinAmort := GetDateFinAmort;    // fonction modifiée pour prime et sbv
  // ajout mbo pour éviter les lignes à 0 en fin d'immoamor
  // recherche fin exercice de l'exercice de fin d'amortissement
  GetDatesExercice(DateFinAmort, DateDebut, DateFin);



  for Lig := 0 to MAX_LIGNEDOT do
  begin
    // fq 19630 - mbo - 16.01.2007
    D2 := 0;
    D3 := 0;
    D4 := 0;
    if TableauDate[Lig] > DateFin then
    begin
      If AmortFisc.Methode <> '' then D2 := AmortFisc.TableauDot[Lig];
      if AmortPRI.Methode <> '' then D3:= AmortPRI.TableauDot[Lig];
      if AmortSBV.Methode <> '' then D4:= AmortSBV.TableauDot[Lig];

      if (AmortEco.TableauDot[lig]=0) and (D2=0) and (D3=0) and (D4=0) then
            TableauDate[Lig] := iDate1900;
    end;
  end;

  Lig := 0;

  while TableauDate[Lig] <> iDate1900 do
  begin
    T := TOB.Create('IMMOAMOR', TAmort, -1);

    T.PutValue('IA_IMMO', CodeImmo);
    T.PutValue('IA_COMPTEIMMO', CompteImmo);
    T.PutValue('IA_NUMEROSEQ', NumSeq);
    T.PutValue('IA_CHANGEAMOR', TypeOpe);

    // TGA 07/04/2006
    T.PutValue('IA_NONDEDUCT',0);
    T.PutValue('IA_CESSIONND',0);
    T.PutValue('IA_MONTANTSBV',0);
    T.PutValue('IA_CESSIONSBV',0);
    T.PutValue('IA_MONTANTDPI',0);
    T.PutValue('IA_CESSIONDPI',0);
    T.PutValue('IA_MONTANTARD',0);
    T.PutValue('IA_CESSIONARD',0);

    if TableauDate[Lig] <> iDate1900 then
    begin
      T.PutValue('IA_DATE', TableauDate[Lig]);
      {Dotation économique}
      Methode := AmortEco;
      DotEco := Methode.TableauDot[Lig];
      T.PutValue('IA_MONTANTECO', DotEco);
      T.PutValue('IA_REINTEGRATION', DotEco * CoeffReintegration);
      T.PutValue('IA_QUOTEPART', DotEco * TauxQuotePart);
      {Dotation partie(s) cédée(s)}
      CessEco := 0.00;
      if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
          CessEco := Methode.TabCession[Lig];

      DotFisc := 0.00;
      DotDerog := 0.00;
      DotReint := 0.00;
      CessFisc := 0.00;

      if AmortFisc.Methode <> '' then
      begin
        Methode := AmortFisc;
        DotFisc := Methode.TableauDot[Lig];

        DotDerog := AmortDerog.TableauDot[Lig];
        DotReint := AmortReint.TableauDot[Lig];

        if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
        begin
           CessFisc := Methode.TabCession[Lig];
        end;
      end;

      {Dotation fiscale}
      T.PutValue('IA_MONTANTFISCAL', DotFisc);
      {Dotation Dérogatoire}
      T.PutValue('IA_MONTANTDEROG', DotDerog);
      // ajout pour chantier fiscal   réintégration extra comptable
      T.PutValue('IA_NONDEDUCT', DotReint);

      {Dotation partie(s) cédée(s)}
      T.PutValue('IA_CESSIONECO', CessEco);
      { CA - 20/01/2004 - FQ 13217 - Calcul de la réintégratin fiscale en cas de sortie }
      if (ValeurAchat <> 0) then
        T.PutValue('IA_REINTEGRATION', CessEco * (ValReintegration / ValeurAchat));

      T.PutValue('IA_BASEDEBEXOECO', AmortEco.Base - CumulEco);
      CumulEco := CumulEco + DotEco;
      if ((AmortFisc.Methode) <> '') then
      begin
        T.PutValue('IA_CESSIONFISCAL', CessFisc);
        T.PutValue('IA_BASEDEBEXOFISC', AmortFisc.Base - CumulFisc);
        CumulFisc := CumulFisc + DotFisc;
      end;
      T.PutValue('IA_BASEDEBEXOCESS', AmortEco.BaseDebExoCes[Lig]);

      {Pour stocker les montants prime on utilise les zones dévolues aux subventions}
      CessPri := 0.00;
      if ((AmortPri.Methode) <> '') then
      begin
        Methode := AmortPri;
        DotPri := Methode.TableauDot[Lig];
        if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
            CessPri := Methode.TabCession[Lig];
        T.PutValue('IA_MONTANTSBV', DotPri);
        T.PutValue('IA_CESSIONSBV', CessPri);
      end;

      {Pour stocker les montants subvention on utilise les zones dévolues aux DPI}
      CessSbv := 0.00;
      if ((AmortSbv.Methode) <> '') then
      begin
         Methode := AmortSbv;
         DotSbv := Methode.TableauDot[Lig];
         if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
               CessSbv := Methode.TabCession[Lig];
         T.PutValue('IA_MONTANTDPI', DotSbv);
         T.PutValue('IA_CESSIONDPI', CessSbv);
      end;

    end;
    DatePrec := TableauDate[Lig];
    Lig := Lig + 1;
  end;
  Result := TAmort.InsertDB(nil, True);
finally
  TAmort.Free;
end;
end;

// Suppression physique du plan dans la base
// doit être suivant de l'appel à la fonction Delete pour l'objet Plan

procedure TPlanAmort.Delete;
begin
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="' + CodeImmo +
    '" AND IA_NUMEROSEQ=' + IntToStr(NumSeq));
end;

// Récupération du Plan correspondant au CodeImmo et à NumPlan

function TPlanAmort.Recupere(tCodeImmo, tNumPlan: string): boolean;
var
  QPlan: TQuery;
begin
  result := false; //EPZ 30/11/00
  QPlan := Open(tCodeImmo, tNumPlan);
  if (QPlan <> nil) and (not QPlan.EOF) then
  begin
    QPlan.First;
    Affecte(QPlan);
    result := true; //EPZ 30/11/00
  end;
  Ferme(QPlan);
end;


{***********A.G.L.***********************************************
Auteur  ...... : CAY
Créé le ...... :
Modifié le ... : 19/10/2006
Description .. : Récupération du Plan correspondant au CodeImmo et à une date donnée
Suite......... : mbo - 19/10/2006 - Modif pour les subventions :
Suite ........ : Si la date de l'opération de saisie de la subvention est
Suite ........ : supérieure à la date de calcul : Les antérieurs doivent être à 0
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.RecupereSuivantDate(tDateCalcul: TDateTime; TAmort: TOB =
  nil; TLog: TOB = nil);
var
  QPlan, QLog: TQuery;
  stNumPlan: string;
  bRevisionPlan, AnnuleSBV, AnnulePRI: boolean;
  DureeEco, DureeFisc : integer;
  MethodeEco, MethodeFisc: string;
  i: integer;
  TPlan: TOB;
  CDA, RSB, RPR, ModifBase : boolean;
  EcoReprise, FiscReprise, PRIreprise, SBVreprise : double;
  BasePRI, BaseSBV, EcoBase, FiscBase, OldValeurHT : double;
  debEco : tDateTime;
  TypeOp : string;
  DureeNAM :string;

begin
  bRevisionPlan := false;
  DureeEco := 0;
  DureeFisc := 0;
  EcoReprise := 0;
  FiscReprise := 0;
  CDA := false;
  RPR := false;
  RSB := false;
  PRIreprise := 0;
  SBVreprise := 0;
  BasePRI := 0;
  BaseSBV := 0;
  EcoBase := 0;
  FiscBase := 0;
  OldValeurHT := 0;
  ModifBase := false;
  DebEco := iDate1900;
  DureeNAM := '';
  AnnuleSBV := false;
  AnnulePRI := false;

  //EPZ 20/12/2000  if (tDateCalcul <= DateFinEnCours) then
  if (tDateCalcul <> iDate1900) and (tDateCalcul <= DateFinEnCours) then
  begin
    //mbo - 09.11.06 - car nouvelles opérations - QLog :=
    //  OpenSQL('SELECT IL_IMMO,IL_TYPEOP,IL_DATEOP,IL_PLANACTIFAP,IL_DUREEECO,IL_DUREEFISC,IL_METHODEECO,IL_METHODEFISC FROM IMMOLOG WHERE (IL_IMMO="'
    QLog :=
      OpenSQL('SELECT * FROM IMMOLOG WHERE (IL_IMMO="'
      +  CodeImmo + '" AND IL_DATEOP<="' + USDateTime(DateFinEnCours)
                  + '") ORDER BY IL_DATEOP DESC,IL_PLANACTIFAP DESC', True);
    while not QLog.Eof do
    begin
      // ajout mbo sur opération CDA, CD2, CD3
      TypeOp := QLog.FindField('IL_TYPEOP').AsString;
      if ((TypeOp = 'CDM') or (TypeOp = 'CDA') or (TypeOp = 'CD2') or (TypeOp = 'CD3')) and
        (QLog.FindField('IL_DATEOP').AsDateTime > tDateCalcul) then
      begin
        bRevisionPlan := True;
        MethodeEco := QLog.FindField('IL_METHODEECO').AsString;
        MethodeFisc := QLog.FindField('IL_METHODEFISC').AsString;
        DureeEco := QLog.FindField('IL_DUREEECO').AsInteger;
        DureeFisc := QLog.FindField('IL_DUREEFISC').AsInteger;

        //ajout mbo 09.11.06 pour avoir des antérieurs bons en édition avant date chgt de methode
        if (QLog.FindField('IL_TYPEOP').AsString = 'CDA') then
        begin
          CDA := true;
          EcoReprise := QLog.FindField('IL_CUMANTCESECO').AsFloat;
          FiscReprise := QLog.FindField('IL_CUMANTCESFIS').AsFloat;
          DebEco := QLog.FindField('IL_DATEOPREELLE').AsDateTime;
        end else
        begin
          if MethodeEco = 'NAM' then
             DureeNAM := TRIM (QLog.FindField('IL_CODECB').AsString);
        end;
      end;

      // ajout mbo 13.11.06 - pour modif de base
      if ((TypeOp = 'MBA') or (TypeOp = 'MB2')) and
        (QLog.FindField('IL_DATEOP').AsDateTime > tDateCalcul) then
      begin
        ModifBase := True;
        DureeFisc := QLog.FindField('IL_DUREEFISC').AsInteger;
        MethodeFisc := QLog.FindField('IL_METHODEFISC').AsString;
        EcoBase := QLog.FindField('IL_BASEECOAVMB').AsFloat;
        FiscBase := QLog.FindField('IL_BASEFISCAVMB').AsFloat;
        EcoReprise := QLog.FindField('IL_REPRISEECO').AsFloat;
        FiscReprise := QLog.FindField('IL_REPRISEFISC').AsFloat;
        OldValeurHT := QLog.FindField('IL_MONTANTAVMB').AsFloat;
      end;

      // ajout mbo 10.11.06 - pour subvention (affichage et impression à une date antérieure)
      if (QLog.FindField('IL_TYPEOP').AsString = 'SBV') and
         (QLog.FindField('IL_DATEOP').AsDateTime > tDateCalcul) then
         AnnuleSBV := true;

      // ajout mbo 10.11.06 - pour PRIME (affichage et impression à une date antérieure)
      if (QLog.FindField('IL_TYPEOP').AsString = 'PRI') and
         (QLog.FindField('IL_DATEOP').AsDateTime > tDateCalcul) then
         AnnulePRI := true;


      // ajout mbo 10.11.06 - pour réduction subvention (affichage et impression à une date antérieure)
      if (QLog.FindField('IL_TYPEOP').AsString = 'RSB') and
         (QLog.FindField('IL_DATEOP').AsDateTime > tDateCalcul) then
      begin
         RSB := true;
         BaseSBV := QLog.FindField('IL_MONTANTEXC').AsFloat;

         if TRIM (QLog.FindField('IL_CODECB').AsString) = 'X' then
            SBVreprise := 0
         else
            SBVreprise := QLog.FindField('IL_MONTANTAVMB').AsFloat;
      end;

      // ajout mbo 10.11.06 - pour réduction subvention (affichage et impression à une date antérieure)
      if (QLog.FindField('IL_TYPEOP').AsString = 'RPR') and
         (QLog.FindField('IL_DATEOP').AsDateTime > tDateCalcul) then
      begin
         RPR := true;
         BasePRI := QLog.FindField('IL_MONTANTEXC').AsFloat;
         PRIreprise := QLog.FindField('IL_MONTANTAVMB').AsFloat;
      end;


      //      if QLog.FindField ('IL_DATEOP').AsDateTime < tDateCalcul then
            { On prend en compte les cession du jour du calcul }
      if QLog.FindField('IL_DATEOP').AsDateTime <= tDateCalcul then
      begin
        stNumPlan := QLog.FindField('IL_PLANACTIFAP').AsString;
        break;
      end;
      QLog.Next;
    end;
    Ferme(QLog);
  end
  else
    stNumPlan := IntToStr(NumSeq);

  if bRevisionPlan then
  begin
    // On modifie les durées et méthodes si une révision du plan a eu lieu après la date de calcul
    AmortEco.Duree := DureeEco;
    AmortFisc.Duree := DureeFisc;
    AmortEco.Methode := MethodeEco;
    AmortFisc.Methode := MethodeFisc;

    if (AmortEco.Methode = 'NAM') and (DureeNAM <> '') and (DureeNAM <> '0') then
    begin
      if SBV then
      begin
         AmortSBV.Duree := StrToInt(DureeNAM);
         AmortSBV.Methode := 'LIN';
         SBVDateDeb := DateMiseEnService;
         AmortSBV.Taux := (GetTaux('LIN', SBVDateDeb,SBVDateDeb, AmortSBV.Duree))/100;
     end;

     if PRI then
     begin
         AmortPRI.Duree := StrToInt(DureeNAM);
         AmortPRI.Methode := 'LIN';
         PRIDateDeb := DateMiseEnService;
         AmortPRI.Taux := (GetTaux('LIN', SBVDateDeb,SBVDateDeb, AmortPRI.Duree))/100;
     end;
    end;

    // ajout mbo 09.11.06
    if CDA then
    begin
      AmortEco.Reprise := EcoReprise;
      AmortFisc.Reprise := FiscReprise;
      if DateDebEco = DateDebFis then
      begin
         DateDebEco := DebEco;
         DateDebFis := DebEco;
      end;
    end;

    AmortEco.Taux := (GetTaux(MethodeEco, DateDebEco,DateDebEco, DureeEco))/100;
    AmortFisc.Taux := (GetTaux(MethodeFisc, DateDebFis,DateDebFis, DureeFisc))/100;

  end;

  // ajout mbo pour modif de bases 13.11.06
  if ModifBase then
  begin
    AmortFisc.Duree := DureeFisc;
    AmortFisc.Methode := MethodeFisc;
    AmortEco.Reprise := EcoReprise;
    AmortFisc.Reprise := FiscReprise;
    ValeurHT := OldValeurHT;
    AmortEco.Base := EcoBase;
    AmortFisc.Base := FiscBase;
    AmortFisc.Taux := (GetTaux(MethodeFisc, DateDebFis,DateDebFis, DureeFisc))/100;
  end;

  // ajout pour certaines éditions (notamment suivi des amortissements dérogatoires)
  if AmortFisc.Duree = 0 then
     fiscal := false;

  if AnnulePRI then
  begin
     MNTPrime := 0;
     PRIDuree := 0;
     PRITaux := 0;
     PRImethode := '';
     AmortPRI.Reprise := 0;
     AmortPRI.Base := 0;
  end;

  if AnnuleSBV then
  begin
     MNTSbv := 0;
     SBVTaux := 0;
     SBVDuree := 0;
     SBVMethode := '';
     AmortSBV.Reprise := 0;
     AmortSBV.Base := 0;
  end;

  if RPR then
  begin
     AmortPRI.Reprise := PRIreprise;
     MNTPrime := BasePRI;
     AmortPRI.Base := BasePRI * 2;
  end;

  if RSB then
  begin
     AmortSBV.Reprise := SBVreprise;
     AmortSBV.Base := BaseSBV;
  end;

  TPlan := TOB.Create('', nil, -1);
  if (stNumPlan <> '') and (TAmort <> nil) then
  begin
    for i := 0 to TAmort.Detail.Count - 1 do
    begin
      if TAmort.Detail[i].GetValue('IA_NUMEROSEQ') = StrToInt(stNumPlan) then
        TOB.Create('IMMOAMOR', TPlan, -1).Dupliquer(TAmort.Detail[i], False,
          True, False);
    end;
  end
  else
  begin
    QPlan := Open(CodeImmo, stNumPlan);
    TPlan.LoadDetailDB('IMMOAMOR', '', '', QPlan, True);
    Ferme(QPlan);
  end;
  (*  if (QPlan<>nil) and (not QPlan.EOF) then
    begin
      QPlan.First;
      Affecte(QPlan);
    end;*)
  if TPlan.Detail.Count > 0 then
    Affecte(nil, TPlan);
  TPlan.Free;
end;


// Ouverture de la query contenant le plan correspondant
// au codeimmo et à NumPlan

function TPlanAmort.Open(tCodeImmo, tNumPlan: string; TAmort: TOB = nil):
  TQuery;
var
  QResult: TQuery;
  Requete: string;
begin
  if (tNumPlan = '') then
  begin
    QResult := OpenSQL('SELECT IA_NUMEROSEQ FROM IMMOAMOR WHERE IA_IMMO="' +
      tCodeImmo + '" ORDER BY IA_NUMEROSEQ DESC', TRUE);
    if not QResult.EOF then
      tNumPlan := QResult.FindField('IA_NUMEROSEQ').AsString
    else
      tNumPlan := '0';
    Ferme(QResult);
  end;
  Requete := 'SELECT * FROM IMMOAMOR WHERE IA_IMMO = "' + tCodeImmo + '" AND ' +
    'IA_NUMEROSEQ =' + tNumPlan + ' ORDER BY IA_DATE ASC';
  result := OpenSQL(Requete, true);
end;

// Affectation des données liées à la query contenant le plan courant
// (cf fnc Recupere)

procedure TPlanAmort.Affecte(QPlan: TQuery; TPlan: TOB = nil);
var
  NbCotations: integer;
  i: integer;
begin
  if QPlan <> nil then
  begin
    CodeImmo := QPlan.FindField('IA_IMMO').AsString;
    CompteImmo := QPlan.FindField('IA_COMPTEIMMO').AsString;
    NumSeqPrec := QPlan.FindField('IA_NUMEROSEQ').AsInteger;
    NumSeq := QPlan.FindField('IA_NUMEROSEQ').AsInteger;
    TypeOpe := QPlan.FindField('IA_CHANGEAMOR').AsString;
  end
  else
  begin
    CodeImmo := TPlan.Detail[0].GetValue('IA_IMMO');
    CompteImmo := TPlan.Detail[0].GetValue('IA_COMPTEIMMO');
    NumSeqPrec := TPlan.Detail[0].GetValue('IA_NUMEROSEQ');
    NumSeq := TPlan.Detail[0].GetValue('IA_NUMEROSEQ');
    TypeOpe := TPlan.Detail[0].GetValue('IA_CHANGEAMOR');
  end;
  NbCotations := -1;
  DecalageAffDotations := 0;


  if QPlan <> nil then
  begin
    while not QPlan.EOF do
    begin
      inc(NbCotations);
      TableauDate[NbCotations] := QPlan.FindField('IA_DATE').AsDateTime;
      if (TableauDate[NbCotations] < DateDebutEnCours) then
        inc(DecalageAffDotations);
      AmortEco.TableauDot[NbCotations] :=
        QPlan.FindField('IA_MONTANTECO').AsFloat;
      AmortFisc.TableauDot[NbCotations] :=
        QPlan.FindField('IA_MONTANTFISCAL').AsFloat;
      AmortEco.TabCession[NbCotations] :=
        QPlan.FindField('IA_CESSIONECO').AsFloat;
      AmortFisc.TabCession[NbCotations] :=
        QPlan.FindField('IA_CESSIONFISCAL').AsFloat;

      // ajout mbo 04.09.06
      AmortPri.TableauDot[NbCotations] :=
        QPlan.FindField('IA_MONTANTSBV').AsFloat;

      AmortPri.TabCession[NbCotations] :=
        QPlan.FindField('IA_CESSIONSBV').AsFloat;

      AmortSbv.TableauDot[NbCotations] :=
        QPlan.FindField('IA_MONTANTDPI').AsFloat;

      AmortSbv.TabCession[NbCotations] :=
        QPlan.FindField('IA_CESSIONDPI').AsFloat;

      Fiscal := Fiscal or (AmortFisc.TableauDot[NbCotations] <> 0) or
        (AmortFisc.TabCession[NbCotations] <> 0);

      AmortEco.BaseDebutExo[NbCotations] :=
        QPlan.FindField('IA_BASEDEBEXOECO').AsFloat;
      AmortFisc.BaseDebutExo[NbCotations] :=
        QPlan.FindField('IA_BASEDEBEXOFISC').AsFloat;
      AmortEco.BaseDebExoCes[NbCotations] :=
        QPlan.FindField('IA_BASEDEBEXOCESS').AsFloat;
      AmortFisc.BaseDebExoCes[NbCotations] :=
        QPlan.FindField('IA_BASEDEBEXOCESS').AsFloat;

      // ajout pour chantier fiscal
      {AmortDerog.TableauDot[NbCotations] :=
        QPlan.FindField('IA_MONTANTDEROG').AsFloat;

      AmortReint.TableauDot[NbCotations] :=
        QPlan.FindField('IA_NONDEDUCT').AsFloat;

      AmortReint.TabCession[NbCotations] :=
        QPlan.FindField('IA_CESSIONND').AsFloat;
      }

      QPlan.Next;
    end;
  end
  else
  begin
    for i := 0 to TPlan.Detail.Count - 1 do
    begin
      inc(NbCotations);
      TableauDate[NbCotations] := TPlan.Detail[i].GetValue('IA_DATE');
      if (TableauDate[NbCotations] < DateDebutEnCours) then
        inc(DecalageAffDotations);
      AmortEco.TableauDot[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_MONTANTECO');
      AmortFisc.TableauDot[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_MONTANTFISCAL');
      AmortEco.TabCession[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_CESSIONECO');
      AmortFisc.TabCession[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_CESSIONFISCAL');

      Fiscal := Fiscal or (AmortFisc.TableauDot[NbCotations] <> 0) or
        (AmortFisc.TabCession[NbCotations] <> 0);

      AmortEco.BaseDebutExo[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_BASEDEBEXOECO');
      AmortFisc.BaseDebutExo[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_BASEDEBEXOFISC');
      AmortEco.BaseDebExoCes[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_BASEDEBEXOCESS');
      AmortFisc.BaseDebExoCes[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_BASEDEBEXOCESS');

      // ajout mbo 04.09.06
      AmortPri.TableauDot[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_MONTANTSBV');

      AmortPri.TabCession[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_CESSIONSBV');

      AmortSbv.TableauDot[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_MONTANTDPI');

      AmortSbv.TabCession[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_CESSIONDPI');

      // ajout mbo pour chantier fiscal
      {AmortDerog.TableauDot[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_MONTANTDEROG');

      AmortReint.TableauDot[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_NONDEDUCT');

      AmortReint.TabCession[NbCotations] :=
        TPlan.Detail[i].GetValue('IA_CESSIONND');}

    end;
  end;
  CalculDateFinAmortissement(AmortEco);
  CalculDateFinAmortissement(AmortFisc);

  if PRI then
     CalculDateFinAmortissement(AmortPri);

  if SBV then
        CalculDateFinAmortissement(AmortSbv);

  // ajout pour chantier fiscal
  // alimentation des tableau dotation dérogatoire et reprise en recalculant
  Calcul_Derog_Reint;

end;

// Recalcul le plan d'amortissement - suite maj Query

procedure TPlanAmort.Recalcul(Q: TQuery);
var
  i: integer;
begin
  for i := 0 to MAX_LIGNEDOT do
    TableauDate[i] := iDate1900;
  InitMethodes;
  Calcul(Q, iDate1900);
end;

//Calcul du plan d'amortissement

procedure TPlanAmort.Calcul(Q: TQuery; tDateCalcul: TDateTime);
var
  DateDebutExoRef: TDateTime;
  wMeth: string;
begin
  bDotationExercice := True;
  if tDateCalcul = iDate1900 then
    Charge(Q)
  else
    tDateCalcul := PlusMois(tDateCalcul, NbMoisReprise);

  DateCalcul := tDateCalcul;

  { mbo fq 18339

  // FQ 17569 DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEPIECEA').AsDateTime);
  // fq 18333 - 08.06.06
  if Q.FindField('I_METHODEFISC').AsString <> ''  then
  begin
     if Q.FindField('I_DATEDEBECO').AsDateTime > Q.FindField('I_DATEDEBFIS').AsDateTime then
        DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEDEBFIS').AsDateTime)
     else
        DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEDEBECO').AsDateTime);
  end else
  begin
     DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEDEBECO').AsDateTime);
  end;
  // fin fq 18333

  }
  // FQ 18339    le but est de construire le tableau des dates à partir de la + petite des dates
  if Q.FindField('I_DATEPIECEA').AsDateTime < Q.FindField('I_DATEAMORT').AsDateTime then
     DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEPIECEA').AsDateTime)
  else
     DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEAMORT').AsDateTime);
  // fin fq 18339

  ResetTableauDot(AmortEco, DateDebutExoRef);
  CalculExeptionnel(AmortEco, DateDebutExoRef);
  wMeth := Q.FindField('I_METHODEECO').AsString;

  { FQ 17569 mbo
  if wMeth = 'LIN' then
    DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEAMORT').AsDateTime)
  else
    DateDebutExoRef :=
      GetDateDebutExoRef(Q.FindField('I_DATEPIECEA').AsDateTime);
  }

  //fq 18333 - on alimente avec les dates de deb amort pour calcul des dotations
  DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEDEBECO').AsDateTime);

  CalculDotation(AmortEco, DateDebutExoRef, false);

  if Fiscal then
  begin
    wMeth := Q.FindField('I_METHODEFISC').AsString;
    { fq 17569
    if wMeth = 'LIN' then
       DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEAMORT').AsDateTime)
    else
      DateDebutExoRef :=
        GetDateDebutExoRef(Q.FindField('I_DATEPIECEA').AsDateTime);
    }

    DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEDEBFIS').AsDateTime);
    ResetTableauDot(AmortFisc, DateDebutExoRef);

    //ajout pour chantier fiscal
    ResetTableauDot(AmortDerog, DateDebutExoRef);
    ResetTableauDot(AmortReint, DateDebutExoRef);

    // fq 17102 - ajout mbo 29/11/05
    IF EcoVar = false then  // fq 17154 - ajout mbo 6.12.05
       CalculExeptionnel(AmortFisc, DateDebutExoRef);

    CalculDotation(AmortFisc, DateDebutExoRef, false);

    // ajout pour chantier fiscal
    Calcul_Derog_Reint;
  end;

  if PRI then
  begin
    if Fiscal then
    begin
       wMeth := Q.FindField('I_METHODEFISC').AsString;
       DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEDEBFIS').AsDateTime);
       ResetTableauDot(AmortPri, DateDebutExoRef);
    end else
    begin
       if Q.FindField('I_METHODEECO').AsString = 'NAM' then
       begin
         wMeth := 'LIN';
         DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEAMORT').AsDateTime);
         ResetTableauDot(AmortPri, DateDebutExoRef);
       end else
       begin
          wMeth := Q.FindField('I_METHODEECO').AsString;
          DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_DATEDEBECO').AsDateTime);
          ResetTableauDot(AmortPri, DateDebutExoRef);
       end;
    end;
    CalculDotation(AmortPri, DateDebutExoRef, false);
  end;

  if SBV then
  begin
    if Fiscal then
    begin
       wMeth := Q.FindField('I_METHODEFISC').AsString;
       DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_SBVDATE').AsDateTime);
       ResetTableauDot(AmortSbv, DateDebutExoRef);
    end else
    begin
       if Q.FindField('I_METHODEECO').AsString = 'NAM' then
       begin
         wMeth := 'LIN';
         DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_SBVDATE').AsDateTime);
         ResetTableauDot(AmortSbv, DateDebutExoRef);
       end else
       begin
          wMeth := Q.FindField('I_METHODEECO').AsString;
          DateDebutExoRef := GetDateDebutExoRef(Q.FindField('I_SBVDATE').AsDateTime);
          ResetTableauDot(AmortSbv, DateDebutExoRef);
       end;
    end;
    CalculDotation(AmortSbv, DateDebutExoRef, false);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 07/11/2005
Modifié le ... : 07/11/2005
Description .. : ajout d'un paramètre en entrée : clocumuldepre
Suite ........ : clocumldepre = calculé par la clôture
Suite ........ : = reprise maximale de dépréciation
Suite ........ : nécessaire au calcul car si = 0 il faut repartir sur un calcul
Suite ........ : normal et non sur durée restante
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.CalculPourTraitement(TypeTraitement: string; CloCumulDepre : double);
var
  i: integer;
begin
  for i := 0 to MAX_LIGNEDOT do
    AmortEco.BaseDebExoCes[i] := 0.00;
  for i := 0 to MAX_LIGNEDOT do
    AmortFisc.BaseDebExoCes[i] := 0.00;
  for i := 0 to MAX_LIGNEDOT do
    AmortEco.TabCession[i] := 0.00;
  for i := 0 to MAX_LIGNEDOT do
    AmortFisc.TabCession[i] := 0.00;

  for i := 0 to MAX_LIGNEDOT do
    AmortPri.BaseDebExoCes[i] := 0.00;
  for i := 0 to MAX_LIGNEDOT do
    AmortPri.TabCession[i] := 0.00;

  for i := 0 to MAX_LIGNEDOT do
    AmortSbv.BaseDebExoCes[i] := 0.00;
  for i := 0 to MAX_LIGNEDOT do
    AmortSbv.TabCession[i] := 0.00;

  //ajout chantier fiscal
  for i := 0 to MAX_LIGNEDOT do
    AmortDerog.BaseDebExoCes[i] := 0.00;
  for i := 0 to MAX_LIGNEDOT do
    AmortDerog.TabCession[i] := 0.00;

  for i := 0 to MAX_LIGNEDOT do
    AmortReint.BaseDebExoCes[i] := 0.00;
  for i := 0 to MAX_LIGNEDOT do
    AmortReint.TabCession[i] := 0.00;


  TypeOpe := TypeTraitement;

  // ajout mbo 4.11.05
  if TypeOpe = 'CLO' then
  begin
     RepriseDepre := CloCumulDepre;
     CumulDepre := Cumul_depreciation(Codeimmo,true);
  end;

  for i := 0 to MAX_LIGNEDOT do
  begin
    if (TableauDate[i] > DateFinEnCours) then
    begin
      AmortEco.TableauDot[i] := 0.00;
      AmortFisc.TableauDot[i] := 0.00;
      AmortPri.TableauDot[i] := 0.00;
      AmortSbv.TableauDot[i] := 0.00;
      AmortDerog.TableauDot[i] := 0.00;
      AmortReint.TableauDot[i] := 0.00;

    end;
  end;
  bDotationExercice := True;
  CalculDotation(AmortEco, DateFinEnCours + 1, true);

  if Fiscal then
  begin
    CalculDotation(AmortFisc, DateFinEnCours + 1, true);
    Calcul_Derog_Reint;
  end;

  if PRI then
      CalculDotation(AmortPri, DateFinEnCours + 1, true);

  if SBV then
      CalculDotation(AmortSbv, DateFinEnCours + 1, true);
end;

// cette procedure n'est pas utilisée (elle est appelée par la fonction
// opeserie pour cession - cette fonctionnalité n'est pas proposée ds l'appli - mbo - 31.10.07
procedure TPlanAmort.CalculPlanCede(T: TOB; tDateCalcul: TDateTime);
var
  i: integer;
  QEx: TQuery;
begin
  bDotationExercice := True;
  if tDateCalcul = iDate1900 then
    ChargeTOB(T)
  else
    tDateCalcul := PlusMois(tDateCalcul, NbMoisReprise);
  { CA - 17/02/2004 - Si après ajout de la reprise, on ne trouve pas d'exercice,
                      on se positionne sur le premier exercice après la date d'amortissement
                      Ceci corrige le problème pour les exercices qui commencent à une date différente du premier }
  { FQ 15110 - CA - 17/12/2004 - ExecuteSQL remplacé par ExisteSQL pour corriger problème dotation on calculée }
  if (not ExisteSQL('SELECT EX_EXERCICE FROM EXERCICE WHERE EX_DATEDEBUT<="' +
    USDateTime(tDateCalcul) +
    '" AND EX_DATEFIN>="' + USDateTime(tDateCalcul) + '"')) then
  begin
    QEx := OpenSQL('SELECT EX_DATEDEBUT FROM EXERCICE WHERE EX_DATEDEBUT>="' +
      USDateTime(tDateCalcul) + '" ORDER BY EX_DATEDEBUT ASC', True);
    try
      if not QEx.Eof then
        tDateCalcul := QEx.FindField('EX_DATEDEBUT').AsDateTime;
    finally
      Ferme(QEx);
    end;
  end;

  DateCalcul := tDateCalcul;
  CalculExeptionnel(AmortEco, DateDebutEnCours);

  for i := 0 to MAX_LIGNEDOT do
    AmortEco.BaseDebExoCes[i] := 0.00;

  CalculDotation(AmortEco, tDateCalcul, false); { dotation économique }
  BasculeBaseCession(AmortEco);


  if Fiscal and EcoVar = False then    // ajout du test sur ecovar mbo le 16.02.06
    // fq 17102 - ajout mbo 29/11/05
    CalculExeptionnel(AmortFisc, DateDebutEnCours);

  CalculDotation(AmortFisc, tDateCalcul, false); { dotation fiscale }
  Calcul_Derog_Reint;  {calcul dérogatoire ou réintégration extra comptable}

  CalculDotation(AmortPri, tDateCalcul, false); { dotation PRIME }
  CalculDotation(AmortSbv, tDateCalcul, false); { dotation SUBVENTION }

end;

procedure TPlanAmort.CalculExeptionnel(MethodeAmort: TMethodeAmort; tDateCalcul:
  TDateTime);
var
  Dotation, dExcEco: double;
  DateDebutEx, DateFinEx: TDateTime;
  NumDot: integer;
  RestePeriode: integer;
  PremMois, PremAnnee, NbMois: Word;
begin
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;

  RestePeriode := GetNombrePeriodeMois(MethodeAmort.Duree - NbMoisReprise, 12);
  NumDot := 0;

  while (TableauDate[NumDot] <> iDate1900) and (TableauDate[NumDot] <
    tDateCalcul) do
    NumDot := NumDot + 1;

  while (RestePeriode > 0) do
  begin
    GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
    dExcEco := 0.00;
    TabTypeExc[NumDot] := ' ';
    if (DateFinEx >= DateDebutEnCours) and (DateFinEx <= DateFinEnCours) then
    begin
      GetElementExceptionnel(Dotation, dExcEco, TabTypeExc[NumDot]);
      dExcEco := dExcEco + MethodeAmort.Revision;
    end;
    MethodeAmort.TableauDot[NumDot] := Arrondi(dExcEco, V_PGI.OkDecV);
    TableauDate[NumDot] := DateFinEx;
    inc(NumDot);
    NOMBREMOIS(DateDebutEx, DateFinEx, PremMois, PremAnnee, NbMois);
    if NbMois >= 12 then
      dec(RestePeriode);
  end;
end;

procedure TPlanAmort.CalculDotation(MethodeAmort: TMethodeAmort; tDateCalcul:
  TDateTime; bCloture: boolean);
var i : integer;
begin

(*  if (TypeAmortCession = 'SAN') then
    exit
  else*)
    CalculDateFinAmortissement(MethodeAmort);

  // impact antérieur dépréciation - mbo 02.06
  if (MethodeAmort.Base <= 0.00) or ((MethodeAmort.Reprise + MethodeAmort.RepriseDep) >= MethodeAmort.Base)
    then
  begin
    ResetTableauDot(MethodeAmort, idate1900)
  end
  else
  begin
    if MethodeAmort.Methode = 'LIN' then
      CalculLineaire(MethodeAmort, tDateCalcul, bCloture)
    else if MethodeAmort.Methode = 'DEG' then
      CalculDegressif(MethodeAmort, tDateCalcul, bCloture)
    else if MethodeAmort.Methode = 'VAR' then
      CalculVariable(MethodeAmort, tDateCalcul, bCloture);
  end;

  if (TypeAmortCession = 'SAN') then
  begin
    for i := 0 to MAX_LIGNEDOT do
    begin
      if (TableauDate[i] >= DateFinEnCours) then
      begin
        { On force à 0 l'amortissement de l'année de la sortie }
        MethodeAmort.TableauDot[i] := 0;
        break;
      end;
    end;
  end;
end;

procedure TPlanAmort.CalculReprises;
begin
  CalculDateFinAmortReprise(AmortEco);
  CalculReprise(AmortEco);

  CalculDateFinAmortReprise(AmortFisc);
  CalculReprise(AmortFisc);

  CalculDateFinAmortReprise(AmortPri);
  CalculReprise(AmortPri);

  CalculDateFinAmortReprise(AmortSbv);
  CalculReprise(AmortSbv);

  //ajout pour fq 17512 - mbo
  if (fiscal) then
     CalculReprise_Derog_Reint;

end;

procedure TPlanAmort.CalculReprise(MethodeAmort: TMethodeAmort);
begin
  MethodeAmort.Reprise := 0.00;
  if GetDateDebutAmort(MethodeAmort) > DateDebutEnCours then
    exit;
  if MethodeAmort.Methode = 'LIN' then
    CalculRepriseLineaire(MethodeAmort, GetDateDebutAmort(MethodeAmort),
      DateDebutEnCours)
  else if MethodeAmort.Methode = 'DEG' then
    CalculRepriseDegressif(MethodeAmort, DateDebutEnCours)
  else if MethodeAmort.Methode = 'VAR' then
    CalculRepriseVariable(MethodeAmort, DateDebutEnCours);
end;

//-----------------------------------------------------------------------------------
procedure TPlanAmort.CalculLineaire(MethodeAmort: TMethodeAmort; tDateCalcul:
  TDateTime; bCloture: boolean);
var
  SommeDot, Vnc, Dotation, Cumul, BaseCalcul, Base, Taux, dExcEco: double;
  DateRef, DateDebutEx, DateFinEx: TDateTime;
  DateFin, Datedebut, dateFinAmort, DateDebAmort: TDateTime;
  NumDot, NumDotExc : integer;
  // ajout mbo crc2002-10
  NbjourTotal, NbJourPeriode : integer;
  CalcDeprec : boolean;
  //
  PremMois, PremAnnee, NbMois: Word;
  bForceCalcul: boolean;
  DateDeb : TDateTime;    // fq 17569 date debut amort = i_datedebeco ou i_datedebfis

  DateDebutExSBV, DateFinExSBV : tDateTime;
  savDot:double;

  VncRest : boolean; // chantier fiscal

begin
  if MethodeAmort.PlanUtil = 'FISC' then
     DateDeb := DateDebFis
  else if MethodeAmort.PlanUtil = 'ECO' then
     DateDeb := DateDebEco
  else if MethodeAmort.PlanUtil = 'PRI' then
      DateDeb := PRIDateDeb
  else
      DateDeb := SBVDateDeb;


  bForceCalcul := False;
  Base := MethodeAmort.Base;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  Taux := MethodeAmort.Taux;
  dExcEco := 0.00;
  NumDotExc := 0;
  NumDot := 0;
  SommeDot := 0.00;
  Cumul := 0.00;
  SavDot := 0.00;
  DateDebutExSBV := iDate1900;
  DateFinExSBV := iDate1900;

{--------------------------------------------------------------------------
  Ajout pour chantier fiscal }
  If MethodeAmort.PlanUtil = 'ECO' then
     VncRest := VncRestEco
  else
     VncRest := VncRestFisc;


  { ---------------------------------------------------------------------------------------
   1ère étape : on calcule les antérieurs en prenant en compte :
  les antérieurs saisis + antérieur dépréciation saisi + les dotations des ex précédents
        Reprise                   RepriseDep                  tableau des dot < exo en cours
}
  // impact antérieur dépréciation -
  {pour les subventions, les antérieurs saisis sont pris en compte sur la première dotation calculée}
  if (MethodeAmort.ReprisePrem = false) then
     Cumul := MethodeAmort.Reprise + MethodeAmort.RepriseDep;  // ajout mbo 02.06;

  if ImmoCloture then
       MethodeAmort.TraiteDotationCloture(TableauDate, NumDot, SommeDot, bCloture);

  Cumul := (Cumul + SommeDot);
  Vnc := (Base - Cumul);
  if (Arrondi(Vnc, V_PGI.OkDecV)) < 0 then
    { cas des VNC supérieures à la valeur HT - solution temporaire pour le recalcul des plans }
  begin
    Vnc := 0;
    Cumul := Base;
    Dec(NumDot, 1);
    bForceCalcul := True;
  end;
  //Pour gÚrer l'anterioritÚ non-modifiable
  while (not bForceCalcul) and (TableauDate[NumDot] <> iDate1900) and
    (TableauDate[NumDot] < tDateCalcul) do
  begin
    Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
    inc(NumDot);
    Vnc := Base - Cumul;
    if (Arrondi(Vnc, V_PGI.OkDecV) < 0) then
      { cas des VNC supÚrieures Ó la valeur HT - solution temporaire pour le recalcul des plans }
    begin
      Vnc := 0;
      Cumul := Base;
      Dec(NumDot, 1);
      bForceCalcul := True;
    end;
  end;
// fin de la première étape --------------------------------------------------------------

// avant calcul de la dotation on repère les dates de l'exercice de la subvention
// pour cumuler les antérieurs avec la dotation de cet exercice là
If MethodeAmort.PlanUtil = 'SBV' then
  GetDatesExercice(SBVDateSBV, DateDebutExSbv, DateFinExSbV);


// 2ème étape : calcul de la dotation de l'exercice en cours et suivants
{ CAS NUMERO 1 - cas normal
 on calcule sur le taux si on traite le plan fiscal avec un plan éco = var
 ou   pas de dépréciation en cours et maxi reprise = 0 et pas de calcul plan sur VNC
}

if ((DureeRest = false) and (CumulDepre = 0) and (VncRest = false))
    OR ((EcoVar = true) and (methodeAmort.planUtil = 'FISC'))
    then
begin
  // calcul sur le taux
  while (Arrondi(Vnc, V_PGI.OkDecV) > 0) do
  begin
    GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
    DateRef := DateDebutEx;
    // fq 17569 if (DateMiseEnService >= DateDebutEx) and (DateMiseEnService <= DateFinEx)
    if (DateDeb >= DateDebutEx) and (DateDeb <= DateFinEx) then
        DateRef := DateDeb;
    if (((MethodeAmort.DateFinAmort >= DateDebutEx) and (MethodeAmort.DateFinAmort <= DateFinEx))  // Fin d'amortissement sur exercice en cours
      or  (MethodeAmort.DateFinAmort < DateDebutEx)) // Fin d'amortissement théorique antérieur à la date de début d'exercice - FQ 15477 - CA - 16/03/2005
      then
      Dotation := Vnc
    else
    begin
       BaseCalcul := Base;
       NOMBREMOIS(DateDebutEx, DateFinEx, PremMois, PremAnnee, NbMois);
       if (not bDotationExercice) and (nbMois <> 12) then
         BaseCalcul := Base * (nbMois / 12);
       Dotation := GetDotationLinApres(DateDebutEx, DateFinEx, DateRef,
         iDate1900, BaseCalcul * Taux, bDotationExercice);
    end;

    // ajout mbo fq 18339
    if DateDeb > DateFinex then
       Dotation := 0;

    if MethodeAmort.PlanUtil <> 'SBV' then
    begin
       if TabTypeExc[NumDot] <> 'SDO' then
       begin
         if SortiePlusExcep then
           // Si except+sortie, on ne prend en compte l'excep qu'Ó la fin.
         begin
           dExcEco := MethodeAmort.TableauDot[NumDot];
           if dExcEco <> 0 then
             NumDotExc := NumDot;
           MethodeAmort.TableauDot[NumDot] := Arrondi(Dotation, V_PGI.OkDecV);
         end
         else
           MethodeAmort.TableauDot[NumDot] :=
             Arrondi(MethodeAmort.TableauDot[NumDot] + Dotation, V_PGI.OkDecV);
       end;
    end else
    begin
       if DateFinEx < DateFinExSBV then
       begin
          MethodeAmort.TableauDot[NumDot] := 0;
          savDot := savDot + Dotation;
       end else
       begin
          MethodeAmort.TableauDot[NumDot]:=Dotation;

          if (MethodeAmort.ReprisePrem) and (DateFinEx = DateFinExSBV) then
              MethodeAmort.TableauDot[NumDot] :=
                    MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise + savDot
          else if (DateFinEx = DateFinExSBV) then
              MethodeAmort.TableauDot[NumDot] :=
                    MethodeAmort.TableauDot[NumDot] + savDot;

       end;
    end;

    if (SortiePlusExcep) and (NumDot = NumDotExc) and (EcoVar = false)
       and (MethodeAmort.PlanUtil <> 'PRI') and (MethodeAmort.PlanUtil <> 'SBV') then
      MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] +
                                         dExcEco;

    //if (MethodeAmort.ReprisePrem) and (NumDot = 0) and (DateDeb <= DateFinex) then
    //   MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise;

    if MethodeAmort.TableauDot[NumDot] > Vnc then
           MethodeAmort.TableauDot[NumDot] := Vnc;


    Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
    TableauDate[NumDot] := DateFinEx;
    inc(NumDot);
    Vnc := Base - Cumul;
  end;
  while MethodeAmort.TableauDot[NumDot] <> 0.00 do
  begin
    MethodeAmort.TableauDot[NumDot] := 0.00;
    inc(NumDot);
  end;
end
else

{ cas numéro 2
  calcul du plan futur sur la VNC (VncRest = true)
 OU maximum reprise (cumuldepre) <> 0
  le calcul sur la durée s'applique à l'exercice en cours et les suivants
}
 if (VncRest = true) or (CumulDepre <> 0) then
 begin
  // inutile - correction conseil de compil - mbo 25.08.06 CalcDeprec := true;

  while (Arrondi(Vnc, V_PGI.OkDecV) > 0) do
  begin
    GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
    DateDebut := DateDebutEx;
    // ajout mbo pour correction avertissement à la compil
    DateDebAmort := DateDebutEx;

    // fq 17569 if (DateMiseEnService >= DateDebutEx) and (DateMiseEnService <= DateFinEx) then
    if (DateDeb >= DateDebutEx) and (DateDeb <= DateFinEx) then
        DateDebAmort := DateDeb;

    if (((MethodeAmort.DateFinAmort >= DateDebutEx) and (MethodeAmort.DateFinAmort <= DateFinEx))  // Fin d'amortissement sur exercice en cours
      or  (MethodeAmort.DateFinAmort < DateDebutEx)) // Fin d'amortissement théorique antérieur à la date de début d'exercice - FQ 15477 - CA - 16/03/2005
      then
      Dotation := Vnc
    else
    begin
      DateFinAmort := MethodeAmort.DateFinAmort;
      DateFin := DateFinEx;

      TraiteParamDateFinAmort(DateDebAmort,DateFinAmort,DateDebut,DateFin,TDateCalcul); // fin amort ou cessio

      NbJourPeriode := NombreJour360(DateDebut, DateFin);
      NbJourTotal:= NombreJour360(DateDebut, DateFinAmort);
      Dotation := Arrondi(((Vnc * NbJourPeriode)/NbJourTotal),V_PGI.OkDecV);
    end;

    // ajout mbo fq 18339
    if DateDeb > DateFinex then
       Dotation := 0;

    if TabTypeExc[NumDot] <> 'SDO' then
    begin
        if SortiePlusExcep then
        // Si except+sortie, on ne prend en compte l'excep qu'à la fin.
        begin
          dExcEco := MethodeAmort.TableauDot[NumDot];
          if dExcEco <> 0 then
            NumDotExc := NumDot;
          MethodeAmort.TableauDot[NumDot] := Arrondi(Dotation, V_PGI.OkDecV);
        end
        else
          MethodeAmort.TableauDot[NumDot] :=
            Arrondi(MethodeAmort.TableauDot[NumDot] + Dotation, V_PGI.OkDecV);
    end;

    if MethodeAmort.PlanUtil <> 'SBV' then
    begin
      if (SortiePlusExcep) and (NumDot = NumDotExc) and (MethodeAmort.PlanUtil <> 'PRI') then
         MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] + dExcEco;
    end else
    begin
      if DateFinEx < DateFinExSBV then
      begin
          MethodeAmort.TableauDot[NumDot] := 0;
          savDot := savDot + Dotation;
      end else
      begin
          MethodeAmort.TableauDot[NumDot]:=Dotation;

          if (MethodeAmort.ReprisePrem) and (DateFinEx = DateFinExSBV) then
              MethodeAmort.TableauDot[NumDot] :=
                    MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise + savDot
          else if (DateFinEx = DateFinExSBV) then
              MethodeAmort.TableauDot[NumDot] :=
                    MethodeAmort.TableauDot[NumDot] + savDot;

      end;
    end;

    if MethodeAmort.TableauDot[NumDot] > Vnc then
      MethodeAmort.TableauDot[NumDot] := Vnc;


    // if (MethodeAmort.ReprisePrem) and (NumDot = 0) and (DateDeb <= Datefinex) then
    //   MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise;

    Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
    TableauDate[NumDot] := DateFinEx;
    inc(NumDot);
    Vnc := Base - Cumul;
  end;

  while MethodeAmort.TableauDot[NumDot] <> 0.00 do
  begin
    MethodeAmort.TableauDot[NumDot] := 0.00;
    inc(NumDot);
  end;

end
else

{Cas numéro 3 :
 saisie d'une dépréciation sur l'exercice
 calcul sur le taux pour l'exercice en cours
 et calcul sur la durée pour exercices suivants

}
 if (DureeRest = true) then
 begin
  CalcDeprec := true;
  while (Arrondi(Vnc, V_PGI.OkDecV) > 0) do
  begin
    GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
    DateRef := DateDebutEx;

    // fq 17569 if (DateMiseEnService >= DateDebutEx) and (DateMiseEnService <= DateFinEx)
    if (DateDeb >= DateDebutEx) and (DateDeb <= DateFinEx) then
        DateRef := DateDeb;

    // mbo 23.03.06 ajout pour correction avertissement en compil
    DateDebAmort := DateRef;

    if (((MethodeAmort.DateFinAmort >= DateDebutEx) and (MethodeAmort.DateFinAmort <= DateFinEx))  // Fin d'amortissement sur exercice en cours
      or  (MethodeAmort.DateFinAmort < DateDebutEx)) // Fin d'amortissement théorique antérieur à la date de début d'exercice - FQ 15477 - CA - 16/03/2005
      then
      Dotation := Vnc
    else
    begin
      //
      if CalcDeprec = true then   // calcul sur taux pour exercice en cours
      begin
         BaseCalcul := Base;
         NOMBREMOIS(DateDebutEx, DateFinEx, PremMois, PremAnnee, NbMois);
         if (not bDotationExercice) and (nbMois <> 12) then
           BaseCalcul := Base * (nbMois / 12);
         Dotation := GetDotationLinApres(DateDebutEx, DateFinEx, DateRef,
         iDate1900, BaseCalcul * Taux, bDotationExercice);
         CalcDeprec:= false;
      end
      else
      // on traite les exercices suivants
      begin
      DateFinAmort := MethodeAmort.DateFinAmort;
      DateDebut := DateDebutEx;
      DateFin := DateFinEx;

      TraiteParamDateFinAmort(DateDebAmort,DateFinAmort,DateDebut,DateFin,TDateCalcul); // fin amort ou cessio

      NbJourPeriode := NombreJour360(DateDebutEx, DateFin);
      NbJourTotal:= NombreJour360(DateDebut, DateFinAmort);
      Dotation := Arrondi(((Vnc * NbJourPeriode)/NbJourTotal),V_PGI.OkDecV);
      end;
    end;

    // ajout mbo fq 18339
    if DateDeb > DateFinex then
       Dotation := 0;

    if TabTypeExc[NumDot] <> 'SDO' then
    begin
      if SortiePlusExcep then
        // Si except+sortie, on ne prend en compte l'excep qu'à la fin.
      begin
        dExcEco := MethodeAmort.TableauDot[NumDot];
        if dExcEco <> 0 then
          NumDotExc := NumDot;
        MethodeAmort.TableauDot[NumDot] := Arrondi(Dotation, V_PGI.OkDecV);
      end
      else
        MethodeAmort.TableauDot[NumDot] :=
          Arrondi(MethodeAmort.TableauDot[NumDot] + Dotation, V_PGI.OkDecV);
    end;


    if (SortiePlusExcep) and (NumDot = NumDotExc) and (MethodeAmort.PlanUtil <> 'PRI') and
       (MethodeAmort.PlanUtil <> 'SBV') then
        MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] + dExcEco;

    //if (MethodeAmort.ReprisePrem) and (NumDot = 0) and (DateDeb <= DateFinex) then
    //   MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise;

    if MethodeAmort.PlanUtil = 'SBV' then
    begin
      if DateFinEx < DateFinExSBV then
      begin
          MethodeAmort.TableauDot[NumDot] := 0;
          savDot := savDot + Dotation;
      end else
      begin
          MethodeAmort.TableauDot[NumDot]:=Dotation;

          if (MethodeAmort.ReprisePrem) and (DateFinEx = DateFinExSBV) then
              MethodeAmort.TableauDot[NumDot] :=
                    MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise + savDot
          else if (DateFinEx = DateFinExSBV) then
              MethodeAmort.TableauDot[NumDot] :=
                    MethodeAmort.TableauDot[NumDot] + savDot;
      end;
    end;

    if MethodeAmort.TableauDot[NumDot] > Vnc then
       MethodeAmort.TableauDot[NumDot] := Vnc;


    Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
    TableauDate[NumDot] := DateFinEx;
    inc(NumDot);
    Vnc := Base - Cumul;
  end;
  while MethodeAmort.TableauDot[NumDot] <> 0.00 do
  begin
    MethodeAmort.TableauDot[NumDot] := 0.00;
    inc(NumDot);
  end;
 end;
end;

//===================================================================================================
{ mbo 25.10.06 modification de la procédure pour intégrer le calcul de la dotation
 d'amortissement subvention qui est particulière :
 forumle = (dotation éco ou fiscale) * (base subvention / base éco ou fiscale)
 il faut donc que le calcul de la dot éco ou fiscale soit toujours fait avant celle de la sbv }

procedure TPlanAmort.CalculDegressif(MethodeAmort: TMethodeAmort; tDateCalcul:
  TDateTime; bCloture: boolean);
var
  EvalLineaire, Prorata, Dotation, Cumul, Base, Taux: double;
  SommeDot: double;
  DateRef, DateDebutEx, DateFinEx, wDate: TDateTime;
  NumDot, RestePeriode: integer;
  DateDebut : TdateTime;
  DejaFait, Afaire : boolean;
  DateDebutExSBV, DateFinExSBV : tDateTime;
  savDot:double;

  //PremMois,PremAnnee,NbMois : Word;
begin
  DejaFait := false;
  aFaire := false;
  Base := MethodeAmort.Base;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  Taux := MethodeAmort.Taux;
  NumDot := 0;
  SommeDot := 0.00;
  Cumul := 0.00;
  SavDot := 0.00;
  DateDebutExSBV := iDate1900;
  DateFinExSBV := iDate1900;
  //DateDebut := iDate1900;

  // fq 17569 ajout mbo 27.04.06
  if MethodeAmort.planUtil = 'FISC' then
     DateDebut := DateDebFis
  else if (MethodeAmort.PlanUtil = 'ECO') then
     DateDebut := DateDebEco
  else if (MethodeAMort.PlanUtil = 'PRI') then
      DateDebut := PRIDateDeb
  else
      DateDebut := SBVDateDeb;

  // impact antérieur dépréciation
  if MethodeAmort.ReprisePrem = false then
     Cumul := MethodeAmort.Reprise + MethodeAmort.RepriseDep;  // ajout mbo 02.06;

  // mbo - 03.10.06 : la propriété création est à true si methode PRI ou SBV et en saisie de PRi ou SBV
  if (ImmoCloture) and (MethodeAmort.creation = false)  then
    MethodeAmort.TraiteDotationCloture(TableauDate, NumDot, SommeDot, bCloture);

  Cumul := Cumul + SommeDot;
  Vnc := Base - Cumul;

  //Calcul du nombre de periodes restantes
  wDate := tDateCalcul;
  RestePeriode := 0;
  while wDate <= MethodeAmort.DateFinAmort do
  begin
    inc(RestePeriode);
    GetDatesExercice(wDate, DateDebutEx, DateFinEx);
    wDate := DateFinEx + 1;
  end;

  // mbo 15.06.06 - fq 15234 - si on a ajouté une période lors de la création du tableauDate
  // on fait -1 sur le nb de période de traitement
  // et on rajoutera une période après avoir traité l'exercice < 12 mois
  If AjoutPeriode then
  begin
     if RestePeriode > 0 then
        RestePeriode := RestePeriode - 1;
  end;

  //Pour gérer l'anteriorité non-modifiable
  while (TableauDate[NumDot] <> iDate1900) and (TableauDate[NumDot] <
    tDateCalcul) do
  begin
    Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
    inc(NumDot);
    Vnc := Base - Cumul;
  end;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  //DateRef := iDate1900;

  // modif mbo 25.10.06 - calcul de la dotation pour la subvention
  // on repart du tableau dot soit de la méthode éco  (planRef = 1)
  //                          soit de la méthode fisc (PlanRef = 2)

If methodeAmort.PlanUtil = 'SBV' then
begin

   // avant calcul de la dotation on repère les dates de l'exercice de la subvention
   // pour cumuler les antérieurs avec cette la dotation de cet exercice là
   If MethodeAmort.PlanUtil = 'SBV' then
           GetDatesExercice(SBVDateSBV, DateDebutExSbv, DateFinExSbV);

   while Arrondi(Vnc, V_PGI.OkDecV) > 0.00 do
   begin
     Dotation := Vnc;     // pour solder sur la dernière annuité (évite pb d'arrondi)

     if MethodeAmort.DateFinAmort > TableauDate[NumDot] then
     begin

        if MethodeAmort.PlanRef = 2 then
           Dotation := (AmortFisc.TableauDot[NumDot]) * (MethodeAmort.Base / Amortfisc.Base)
        else
           Dotation := (AmortEco.TableauDot[NumDot]) * (MethodeAmort.Base / AmortEco.Base);
     end;

     // on va tester par rapport aux dates du tableaudate pour savoir si la dotation doit être
     // affectée à cet indice du tableau ou à un indice postérieur
     if TableauDate[NumDot] < DateFinExSBV then
     begin
        MethodeAmort.TableauDot[NumDot] := 0;
        savDot := savDot + Dotation;
     end else
     begin
        MethodeAmort.TableauDot[NumDot]:=Dotation;

        if (MethodeAmort.ReprisePrem) and (TableauDate[NumDot] = DateFinExSBV) then
           MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise + savDot
        else if (TableauDate[NumDot] = DateFinExSBV) then
           MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] + savDot;

     end;

     MethodeAmort.TableauDot[NumDot] := Arrondi(MethodeAmort.TableauDot[NumDot], V_PGI.OkDecV);

     // pour pouvoir affecter sur l'indice du tableaudot qui va bien
     {
     if (MethodeAmort.ReprisePrem) then
     begin
       if (NumDot = 0) then
       begin
          if (tDateCalcul <= TableauDate[NumDot]) then
             MethodeAmort.TableauDot[NumDot]:= MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise;
       end else
       begin
          if (tDateCalcul > TableauDate[NumDot-1]) and (tDateCalcul <= TableauDate[NumDot]) then
             MethodeAmort.TableauDot[NumDot]:= MethodeAmort.TableauDot[NumDot] + MethodeAmort.Reprise;
       end;
     end;
     }

     Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
     inc(NumDot);
     Vnc := Base - Cumul;
   end;

end else
begin
  while Arrondi(Vnc, V_PGI.OkDecV) > 0.00 do
  begin
    GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
    DateRef := DateDebutEx;

    //modif mbo le 27.04.2006 - fq 17569 - on prend en cpte la date début d'amortissement
    //if (DateAchat > DateDebutEx) and (DateAchat <= DateFinEx) then
    //  DateRef := DateAchat;
    if (DateDebut > DateDebutEx) and (DateDebut <= DateFinEx) then
        DateRef := DateDebut

    else  // ajout 15.06.06 fq 15234 - on teste si l'exercice que l'on traite est < 12 mois
       if (not DejaFait) then
          if ExoInferieurA12(DateDebutEx)then Afaire := true;

    Dotation := Vnc;

    if MethodeAmort.DateFinAmort > DateFinEx then
    begin
      Dotation := GetDotationDegApres(DateDebutEx, DateFinEx, DateRef, (*Base*)
        Vnc * Taux, Prorata);

      if (RestePeriode > 0) then
      begin
        EvalLineaire := ((Vnc / RestePeriode) * Prorata);
        if (EvalLineaire > Dotation) then
          Dotation := Arrondi(EvalLineaire, V_PGI.OkDecV);
      end;
    end;
    // Calcul du suramortissement le cas échéant
    if SurAmort then
      // fq 17569 remplacement DateAchat par DateDebut
      Dotation := CalculSurAmortissement(Dotation, DateDebut, DateDebutEx,DateFinEx);

    // ajout mbo fq 18339 - pour s'assurer que si l'amortissement n'a pas commencé : la dot = 0
    if DateDebut > DateFinex then
       Dotation := 0;

    if TabTypeExc[NumDot] <> 'SDO' then
    begin
      MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] +
        Dotation;
      MethodeAmort.TableauDot[NumDot] :=
        Arrondi(MethodeAmort.TableauDot[NumDot], V_PGI.OkDecV);
    end;

    Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
    TableauDate[NumDot] := DateFinEx;
    inc(NumDot);
    Dec(RestePeriode);
    Vnc := Base - Cumul;

    // ajout mbo 15.06.06 - fq 15234
    // si on doit ajouter une période, si ce n'est pas déjà
    // et si l'on a détecté que l'exercice que l'on vient de traiter < 12 mois
    // alors on ajoute une période de traitement pour calculer les dotations sur les exo suivants
    // ce traitement ne sera fait qu'une fois
    if (AjoutPeriode) and (not DejaFait) and (Afaire = true) then
      begin
         Inc(RestePeriode);
         AjoutPeriode := false;
         DejaFait := true;
         Afaire := false;
      end;

  end;
end;

  while MethodeAmort.TableauDot[NumDot] <> 0.00 do
  begin
    MethodeAmort.TableauDot[NumDot] := 0.00;
    NumDot := NumDot + 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 25/05/2004
Modifié le ... :   /  /
Description .. : Calcul de dotation pour une méthode d'amortissement
Suite ........ : variable (par unités d'oeuvre)
Mots clefs ... :
*****************************************************************}

{mbo 02.11.05 modif pour depreciation
 procedure TPlanAmort.CalculVariable(MethodeAmort: TMethodeAmort; DateCalcul:
  TDateTime; bCloture : boolean );
  function ProrataUO(NombreUO, TotalUO: double): double;
  begin
//    Result := Arrondi((NombreUO / TotalUO), 9);   // Attention : on met 9 pour avoir un calcul précis
    Result := (NombreUO / TotalUO);
  end;

var
  i, NumDot: integer;
  CumulDotation : double;
begin
  NumDot := 0;
  CumulDotation := MethodeAmort.Reprise;
  for i := 0 to fPlanUO.Detail.Count - 1 do
  begin
    if (bCloture) or (fPlanUO.Detail[i].GetValue('IUO_DATE') >= DateCalcul) then
    begin
      TableauDate[NumDot] := fPlanUO.Detail[i].GetValue('IUO_DATE');
      MethodeAmort.TableauDot[NumDot] := Arrondi (
      MethodeAmort.Base * ProrataUO(fPlanUO.Detail[i].GetValue('IUO_UNITEOEUVRE'), fTotalUO)
      , V_PGI.OkDecV);
      CumulDotation := CumulDotation + MethodeAmort.TableauDot[NumDot];
      Inc(NumDot);
    end;
  end;
  //  Pour que la totalisation soit exacte, on solde la dernière ligne

  if (NumDot > 0) then MethodeAmort.TableauDot[NumDot-1] := Arrondi(MethodeAmort.TableauDot[NumDot-1] + (MethodeAmort.Base-CumulDotation),V_PGI.OkDecV);
  // Raz des dates inutilisées
  while ((TableauDate[NumDot] <> iDate1900) or ((MethodeAmort.TableauDot[NumDot] <> 0.00))) do
  begin
    TableauDate[NumDot] := iDate1900;
    MethodeAmort.TableauDot[NumDot] := 0.00;
    NumDot := NumDot + 1;
  end;
end;
}

// réécriture de la procédure calcul de la méthode variable
procedure TPlanAmort.CalculVariable(MethodeAmort: TMethodeAmort; DateCalcul:
  TDateTime; bCloture : boolean );
  function ProrataUO(NombreUO, TotalUO: double): double;
  begin
//    Result := Arrondi((NombreUO / TotalUO), 9);   // Attention : on met 9 pour avoir un calcul précis
    Result := (NombreUO / TotalUO);
  end;

var
  i, NumDot, ExoEnCours: integer;
  Cumul : double;
  Base : double;
  CumulDotation : double;
  CumulUO : double;
  ResteUO : double;
  Prorata : double;
  VNC : double;
  VncAvant : double;
  bForceCalcul: boolean;
  AjoutExc : double;
  Indice_UO : integer;
  VncRest : boolean;
begin
//VncAvant := 0;
bForceCalcul := false;
Base := MethodeAmort.Base;
CumulUo := 0;
Cumul := 0; //MVG 13/04/2006
//VncRest:= false;

// ajout pour chantier fiscal
VncRest := VncRestEco;

// 1ère étape on recalcule les antérieurs comme si on pouvait saisir avec antérieurs
// antérieurs = antérieurs saisis + antérieurs dépréciation saisis + cumul tableau dot < date de calcul
  NumDot := 0;
  CumulDotation := MethodeAmort.Reprise + MethodeAmort.RepriseDep;  // fq 15274

  // on gère l'antériorité non modifiable pour tenir compte d'un éventuel exceptionnel
  Vnc := (Base - CumulDotation);
  if (Arrondi(Vnc, V_PGI.OkDecV)) < 0 then
    { cas des VNC supérieures à la valeur HT - solution temporaire pour le recalcul des plans }
  begin
    Vnc := 0;
    Cumul := Base;
    Dec(NumDot, 1);
    bForceCalcul := True;
  end;

  // fq 18339 on va recuperer l'indice du fPlanUO correspondant à l'exercice que l'on traite
  indice_UO := 0;
  i:= numdot;

  // fq 18595 - réécriture recherche indice éclatée en 2 parties
  // 1ère partie : pour le cas où date de mise en service > exercice d'achat
  while (not bForceCalcul) and (TableauDate[i] <> iDate1900) do
  begin
     if (TableauDate[i]< fPlanUO.Detail[0].GetValue('IUO_DATE'))
     then
        inc(indice_UO);       // cas d'une mise en service au-delà de l'exercice d'acquisition

     inc(i);
  end;

  // fq 18595 : 2ème partie : cas d'une mise en service antérieure à l'exercice en cours
  i := 0;
  if indice_UO = 0 then
  begin
     while (not BForceCalcul) and (i <= fPlanUO.Detail.Count-1) do
        // ajout pour fq 15274  pour gérer les antérieurs saisis : reprise
     begin
        if fPlanUO.Detail[i].GetValue('IUO_DATE') < TableauDate[0] then
        begin
           dec(indice_UO);
           CumulUO := CumulUO + fPlanUO.Detail[i].GetValue('IUO_UNITEOEUVRE');
        end;
        inc(i);
     end;

  end;

  //Pour gérer l'anteriorité non-modifiable
  while (not bForceCalcul) and (TableauDate[NumDot] <> iDate1900) and
     (TableauDate[NumDot] < DateCalcul) do
  begin
    CumulDotation := (CumulDotation + MethodeAmort.TableauDot[NumDot]);

    if CumulUo < fTotalUO then
      if Numdot >= indice_UO then
         CumulUo := (CumulUO + fPlanUO.Detail[NumDot-indice_UO].GetValue('IUO_UNITEOEUVRE'));

    { modif mbo pour conseil de compil 25.08.06 - prorata non utilisé ds ce cas
    ResteUo := (fTotalUO - CumulUo);
    Vnc := (Base - Cumul);

    if ResteUo = 0.00 then
       prorata := 0
    else
    begin
       if Numdot >= indice_UO then
          prorata := (fPlanUO.Detail[NumDot+1-indice_UO].GetValue('IUO_UNITEOEUVRE')/ResteUO)
       else
          prorata := 0;
    end;
    }

    inc(NumDot);
    Vnc := (Base - CumulDotation);
    if (Arrondi(Vnc, V_PGI.OkDecV) < 0) then
      { cas des VNC supérieures à la valeur HT - solution temporaire pour le recalcul des plans }
    begin
      Vnc := 0;
      Cumul := Base;
      Dec(NumDot, 1);
      bForceCalcul := True;
    end;
  end;
// fin calcul des antérieurs ---------------------------------------------------------------

// 2ème étape : on traite les exercices N et suivants
if ((DureeRest = false) and (CumulDepre = 0) and (VncRest = false)) then
begin
  // on est dans le cas normal
  // on traite exercice N et suivants
  ExoEnCours := Numdot;

  for i := NumDot to fPlanUO.Detail.Count+indice_Uo - 1 do
  begin
    if (bCloture) or (fPlanUO.Detail[i-indice_UO].GetValue('IUO_DATE') >= DateCalcul) then
    begin

      if i >= indice_UO then
         TableauDate[i] := fPlanUO.Detail[i-indice_UO].GetValue('IUO_DATE');

      VncAvant := (Base - CumulDotation);

      //ajout pour l'exceptionnel
      if (i = ExoEnCours) then // fq 15274 and (indice_UO = 0)then
      begin
         if TypeExc = 'RDO' then
            AjoutExc := (MontantExc * -1)
         else
            AjoutExc := MontantExc;
      end else
         AjoutExc := 0;

      if i >= indice_UO then
         MethodeAmort.TableauDot[i] := AjoutExc + Arrondi (
                   MethodeAmort.Base * ProrataUO(fPlanUO.Detail[i-indice_UO].GetValue('IUO_UNITEOEUVRE'), fTotalUO)
                  , V_PGI.OkDecV)
      else
         MethodeAmort.TableauDot[i] := AjoutExc;

      if VncAvant < MethodeAmort.TableauDot[i] then
         MethodeAmort.TableauDot[i] := VncAvant;

      MethodeAmort.TableauDot[i]:= Arrondi(MethodeAmort.TableauDot[i], V_PGI.OkDecV);
      CumulDotation := (CumulDotation + MethodeAmort.TableauDot[i]);
      // ajout pour exceptionnel
      Vnc := Arrondi((Base - CumulDotation), V_PGI.OkDecV);
      if Vnc < 0.00 then MethodeAmort.TableauDot[i]:= VncAvant;
      Inc(NumDot);
    end;
  end;

  //  Pour que la totalisation soit exacte, on solde la dernière ligne
  if (NumDot > 0) then MethodeAmort.TableauDot[NumDot-1] := Arrondi(MethodeAmort.TableauDot[NumDot-1] + (MethodeAmort.Base-CumulDotation),V_PGI.OkDecV);

  // Raz des dates inutilisées
  while ((TableauDate[NumDot] <> iDate1900) or (MethodeAmort.TableauDot[NumDot] <> 0.00)) do
  begin
      TableauDate[NumDot] := iDate1900;
      MethodeAmort.TableauDot[NumDot] := 0.00;
      NumDot := NumDot + 1;
  end;
end
else
  // on traite le cas de calcul sur durée sur N et suivants
if (VncRest = true) or (CumulDepre <> 0) then
begin
 ExoEnCours := Numdot;
 VncAvant := Arrondi((Base - CumulDotation), V_PGI.OkDecV);

  for i := NumDot to fPlanUO.Detail.Count+indice_UO - 1 do
  begin
    if (bCloture) or (fPlanUO.Detail[i-indice_UO].GetValue('IUO_DATE') >= DateCalcul) then
    begin
      if Numdot >= indice_UO then
         TableauDate[NumDot] := fPlanUO.Detail[i-indice_UO].GetValue('IUO_DATE');

      //ajout pour l'exceptionnel  et dépréciation d'actif
      if (bCloture = false) then    // fq 15274 and (indice_UO = 0)then
      begin
         if i = ExoEnCours then
         begin
           if TypeExc = 'RDO' then
              AjoutExc := (MontantExc * -1) + AmortEco.Revision
           else
              AjoutExc := MontantExc + AmortEco.Revision;
         end else
           AjoutExc := 0;
      end else
         AjoutExc := 0;

      ResteUo := (fTotalUO - CumulUo);

      if ResteUo <= 0.00 then
         prorata := 0
      else
         // fq 18339 prorata := (fPlanUO.Detail[NumDot].GetValue('IUO_UNITEOEUVRE')/ResteUO);
         if Numdot >= indice_UO then
            prorata := (fPlanUO.Detail[i-indice_UO].GetValue('IUO_UNITEOEUVRE')/ResteUO)
         else
            prorata := 0;

      MethodeAmort.TableauDot[i] := AjoutExc +
                                         Arrondi ((VncAvant * Prorata), V_PGI.OkDecV);

      if VncAvant < MethodeAmort.TableauDot[i] then
         MethodeAmort.TableauDot[i] := VncAvant;

      MethodeAmort.TableauDot[i]:= Arrondi(MethodeAmort.TableauDot[i], V_PGI.OkDecV);
      CumulDotation := (CumulDotation + MethodeAmort.TableauDot[i]);
      // ajout pour exceptionnel
      Vnc := Arrondi((Base - CumulDotation), V_PGI.OkDecV);
      if Vnc < 0.00 then MethodeAmort.TableauDot[i]:= VncAvant;

      VncAvant := Vnc;

      if CumulUo < fTotalUO then
         //fq 18339 CumulUo := (CumulUO + fPlanUO.Detail[NumDot].GetValue('IUO_UNITEOEUVRE'));
         if Numdot >= indice_UO then
            CumulUo := (CumulUO + fPlanUO.Detail[Numdot-Indice_UO].GetValue('IUO_UNITEOEUVRE'));
      Inc(NumDot);
    end;
  end;

   //  Pour que la totalisation soit exacte, on solde la dernière ligne
  if (NumDot > 0) then MethodeAmort.TableauDot[NumDot-1] := Arrondi(MethodeAmort.TableauDot[NumDot-1] + (MethodeAmort.Base-CumulDotation),V_PGI.OkDecV);

  // Raz des dates inutilisées
  while ((TableauDate[NumDot] <> iDate1900) or (MethodeAmort.TableauDot[NumDot] <> 0.00)) do
  begin
      TableauDate[NumDot] := iDate1900;
      MethodeAmort.TableauDot[NumDot] := 0.00;
      NumDot := NumDot + 1;
  end;
end else

// 3ème cas : on calcule l'exercice N sur le taux et les suivants sur la durée
if (DureeRest = true) and (Vnc > 0.00) then
begin
    ExoEnCours := Numdot;
   // fq 18339 TableauDate[NumDot] := fPlanUO.Detail[NumDot].GetValue('IUO_DATE');
   if Numdot >= indice_UO then
      TableauDate[NumDot] := fPlanUO.Detail[Numdot-indice_UO].GetValue('IUO_DATE');

    // ajout pour l'exceptionnel
    If TypeExc = 'RDO' then
       AjoutExc := MontantExc * -1
    else
       AjoutExc := MontantExc;

   { fq 18339 MethodeAmort.TableauDot[NumDot] := AjoutExc +
        Arrondi (MethodeAmort.Base * ProrataUO(fPlanUO.Detail[NumDot].GetValue('IUO_UNITEOEUVRE'), fTotalUO)
        , V_PGI.OkDecV) + AmortEco.Revision; }
    // fq 15274   if indice_UO <> 0 then
    if Numdot <> ExoEnCours then
       AJoutExc :=0;

    if Numdot >= indice_UO then
        MethodeAmort.TableauDot[NumDot] := AjoutExc +
         Arrondi (MethodeAmort.Base * ProrataUO(fPlanUO.Detail[numdot-Indice_UO].GetValue('IUO_UNITEOEUVRE'), fTotalUO)
         , V_PGI.OkDecV) + AmortEco.Revision
    else
        MethodeAmort.TableauDot[NumDot] := AjoutExc + AmortEco.Revision;

    MethodeAmort.TableauDot[NumDot]:= Arrondi(MethodeAmort.TableauDot[NumDot], V_PGI.OkDecV);
    Cumul := (Cumul + MethodeAmort.TableauDot[NumDot]);


    if CumulUo < fTotalUO then
       // fq 18339 CumulUo := (CumulUO + fPlanUO.Detail[NumDot].GetValue('IUO_UNITEOEUVRE'));
       if Numdot >= indice_UO then
          CumulUo := (CumulUO + fPlanUO.Detail[Numdot-Indice_UO].GetValue('IUO_UNITEOEUVRE'));

    ResteUo := (fTotalUO - CumulUo);
    Vnc := (Base - Cumul);
    if ResteUo = 0.00 then
       prorata := 0
    else
       //fq 18339 prorata := (fPlanUO.Detail[NumDot+1].GetValue('IUO_UNITEOEUVRE')/ResteUO);
       if Numdot >= indice_UO then
          prorata := (fPlanUO.Detail[Numdot-indice_UO+1].GetValue('IUO_UNITEOEUVRE')/ResteUO)
       else
          prorata := 0;

    inc(NumDot);


    if (Arrondi(Vnc, V_PGI.OkDecV) < 0) then
       { cas des VNC supÚrieures Ó la valeur HT - solution temporaire pour le recalcul des plans }
    begin
      Vnc := 0;
      Cumul := Base;
      Dec(NumDot, 1);
      // mbo 24.08.06 pour conseil en compil - bForceCalcul := True;
    end;


  { on traite les exo suivants }
  if (Vnc > 0.00) then
  begin
    for i := NumDot to fPlanUO.Detail.Count+indice_UO - 1 do
    begin
      TableauDate[i] := fPlanUO.Detail[i-indice_UO].GetValue('IUO_DATE');
      MethodeAmort.TableauDot[i] := Arrondi ((Vnc * prorata), V_PGI.OkDecV);

      Cumul := (Cumul + MethodeAmort.TableauDot[i]);
      Vnc := (Vnc - MethodeAmort.TableauDot[i]);
      if CumulUo < fTotalUO then
         CumulUo := (CumulUo + fPlanUO.Detail[i-indice_UO].GetValue('IUO_UNITEOEUVRE'));

      ResteUo := (fTotalUO - CumulUo);
      if Arrondi(ResteUo, V_PGI.OkDecV) > 0 then
         prorata := (fPlanUO.Detail[i+1 - indice_UO].GetValue('IUO_UNITEOEUVRE') / ResteUo)
      else
         prorata := 0;
      inc(NumDot);
    end;

    //  Pour que la totalisation soit exacte, on solde la derniÞre ligne
    if (NumDot > 0) then MethodeAmort.TableauDot[NumDot] := Arrondi((MethodeAmort.TableauDot[NumDot] + (Base-Cumul)),V_PGI.OkDecV);

    // Raz des dates inutilisÚes
    while ((TableauDate[NumDot] <> iDate1900) or (MethodeAmort.TableauDot[NumDot] <> 0.00)) do
    begin
      TableauDate[NumDot] := iDate1900;
      MethodeAmort.TableauDot[NumDot] := 0.00;
      NumDot := NumDot + 1;
    end;
  end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : modif mbo
Créé le ...... : 09/11/2005
Modifié le ... :   /  /
Description .. : prise en compte de la dépréciation d'actif
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.SetTypeOpe(NewTypeOpe: string);
begin
  if (NewTypeOpe = 'RDO') or (NewTypeOpe = 'DOT') or (NewTypeOpe = 'SDO') then
    TypeOpe := 'ELE'
  else if NewTypeOpe = 'DPR' then       // modif mbo 26.10.05
    typeOpe := 'DPR'
  else
    //TypeOpe := 'CDM';    // BTY 11/05 passer l'opération à TypeOpe
    TypeOpe := NewTypeOpe;
end;

procedure TPlanAmort.SetTypeExc(NewTypeExc: string);
begin
  TypeExc := NewTypeExc;
end;

// Récupération BaseFinAmortissement

function TPlanAmort.GetBaseFinAmort: double;
begin
  result := BaseFinAmort;
end;

// Modification BaseFinAmortissement

procedure TPlanAmort.SetBaseAmortFin(MontantOpe: double);
begin
  BaseFinAmort := BaseFinAmort - MontantOpe;
end;

function TPlanAmort.GetDateFinAmortEx(MethAmort: TMethodeAmort): TDateTime;
var
  i: integer;
  RetDate: TDateTime;
begin
  i := 0;
  RetDate := iDate1900;
  while (MethAmort.TableauDot[i] <> 0.00) or (MethAmort.TabCession[i] <> 0.00)
    do
  begin
    RetDate := TableauDate[i];
    i := i + 1;
  end;
  result := RetDate;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 07/11/2005
Modifié le ... : 07/11/2005
Description .. : permet de calculer le nb d'unités d'oeuvre restantes
Suite ........ : au moment de la cloture (appelé par immoclo)
Mots clefs ... :
*****************************************************************}
Function TPLanAmort.ProrataUOReste (DateCalcul:TdateTime): double;
var
  CumulUO : double;
  Prorata : double;
  UOexercice : double;
  i : integer;
begin
  CumulUO:=0; //MVG 13/04/2006
  UOExercice:=0; //MVG 13/04/2006
  for i := 0 to fPlanUO.Detail.Count - 1 do
  begin
    if (fPlanUO.Detail[i].GetValue('IUO_DATE') <= DateCalcul) then
    begin
      CumulUo := cumulUO + fPlanUO.Detail[i].GetValue('IUO_UNITEOEUVRE');
      if i < (fPlanUO.Detail.Count - 1) then
        UOexercice:= fPlanUO.Detail[i+1].GetValue('IUO_UNITEOEUVRE');

    end;
  end;

if CumulUo = fTotalUo then
   prorata := 1
else
   Prorata := UOexercice / (fTotalUo - CumulUO);

// a priori il n'est pas fait d'arrondi Result  := Arrondi (Prorata, V_PGI.OkDecV);
result:=Prorata;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/10/2004
Modifié le ... :   /  /
Description .. : Calcul de la VNC
Suite ........ : - CA - 12/10/2004 - FQ 14664 - Calcul de la VNC par
Suite ........ : rapport à la valeur d'achat et non par rapport à la base
Suite ........ : d'amortissement
Suite......... : modif mbo - 02.06 - prise en compte de l'antérieur dépréciation
Mots clefs ... :
*****************************************************************}
function TPlanAmort.GetVNCAvecMethode(MethAmort: TMethodeAmort; DateRef:
  TDateTime): double;
var
  CumulDot, DotExo: double;
  Indice: integer;
begin
  //09/04/99  CumulDot := MethAmort.GetCumulDotExercice(ImmoCloture,DateRef,TableauDate,false);
  CumulDot := 0.00;
  Indice := 0;
  while (TableauDate[Indice] <> iDate1900) and (TableauDate[Indice] <= DateRef)
    do
  begin
    if MethAmort.TableauDot[Indice] <> 0 then
      DotExo := MethAmort.TableauDot[Indice]
    else
      DotExo := MethAmort.TabCession[Indice];

    CumulDot := CumulDot + DotExo;
    Indice := Indice + 1;
  end;
  //09/04/99
  // BaseCalc := MethAmort.BaseDebutExo[0]; //GetBaseExercice(DateRef,MethAmort);
  //  result := BaseCalc -(CumulDot+MethAmort.Reprise); CA - 22/04/1999
  { CA - 12/10/2004 - FQ 14664 - La VNC se calcule par rapport à la valeur d'achat
    et non par rapport à la base d'amortissement. Ce dernier cas correspond à la valeur
    résiduelle }
//  result := BaseCalc - CumulDot; // car reprise déjà déduite dans basedebutexo

  // impact antérieur dépréciation - mbo 02.06

  // modif mbo pour calcul de la vnc pour la prime d'équipement
  If (MethAmort.PlanUtil = 'ECO') or (MethAmort.PlanUtil = 'FISC') then
    Result := ValeurHT - (CumulDot+MethAmort.Reprise + MethAmort.RepriseDep) + ValeurTVARecuperable-ValeurTVARecuperee
  else if MethAmort.PlanUtil = 'PRI' then
    Result := MethAmort.Base - (CumulDot+MethAmort.Reprise)
  else
  begin
    if MethAmort.ReprisePrem then
       Result := MethAmort.Base - (CumulDot) // car reprise déjà dans la première dotation
    else
       Result := MethAmort.Base - (CumulDot+ MethAmort.Reprise) // cas ou reprise pas ds 1ère dotation
  end;
  //08/04/99 EPZ  result := ValeurHT-(CumulDot+MethAmort.Reprise);
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 18/09/2006
Modifié le ... :   /  /
Description .. : calcul du reste à amortir sur la prime d'investissement lors de
Suite ........ : la cession de l'immo
Mots clefs ... :
*****************************************************************}
function TPlanAmort.GetRestePri(MethAmort: TMethodeAmort; DateRef:
  TDateTime): double;
var
  CumulDot, DotExo: double;
  Indice: integer;
begin
  //09/04/99  CumulDot := MethAmort.GetCumulDotExercice(ImmoCloture,DateRef,TableauDate,false);
  CumulDot := 0.00;
  Indice := 0;
  while (TableauDate[Indice] <> iDate1900) and (TableauDate[Indice] <= DateRef)
    do
  begin
    if MethAmort.TableauDot[Indice] <> 0 then
      DotExo := MethAmort.TableauDot[Indice]
    else
      DotExo := MethAmort.TabCession[Indice];
    CumulDot := CumulDot + DotExo;
    Indice := Indice + 1;
  end;

  // lors de la cession, pour la prime on solde le reste à amortir
  Result := MethAmort.base  - (CumulDot+MethAmort.Reprise);

end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 28/09/2006
Modifié le ... :   /  /
Description .. : calcul du reste à amortir sur la subvention lors de
Suite ........ : la cession de l'immo
Mots clefs ... :
*****************************************************************}
function TPlanAmort.GetResteSbv(MethAmort: TMethodeAmort; DateRef:
  TDateTime): double;
var
  CumulDot, DotExo: double;
  Indice: integer;
begin
  CumulDot := 0.00;
  Indice := 0;
  while (TableauDate[Indice] <> iDate1900) and (TableauDate[Indice] <= DateRef)
    do
  begin
    if MethAmort.TableauDot[Indice] <> 0 then
      DotExo := MethAmort.TableauDot[Indice]
    else
      DotExo := MethAmort.TabCession[Indice];
    CumulDot := CumulDot + DotExo;
    Indice := Indice + 1;
  end;

  // lors de la cession, pour la prime on solde le reste à amortir
  if not(MethAmort.ReprisePrem) then
     Result := MethAmort.Base - CumulDot - MethAmort.reprise
  else
     Result := MethAmort.base  - CumulDot;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/10/2004
Modifié le ... :   /  /
Description .. : Calcul de la VNC de la partié cédée
Suite ........ : - CA - 12/10/2004 - La VNC se calcul par rapport à la
Suite ........ : valeur HT et non par rapport à la base d'amortissement
Mots clefs ... :
*****************************************************************}
function TPlanAmort.GetVNCCEdee(MethAmort: TMethodeAmort; DateRef: TDateTime):
  double;
var
  BaseCalc, CumulDot: double;
  Indice: integer;
begin
  CumulDot := 0.00;
  Indice := 0;
  BaseCalc := 0;
  while (TableauDate[Indice] <> iDate1900) and (TableauDate[Indice] <= DateRef)
    do
  begin
    CumulDot := CumulDot + MethAmort.TabCession[Indice];
    BaseCalc := MethAmort.BaseDebExoCes[Indice];
    Indice := Indice + 1;
  end;
  // result := BaseCalc - CumulDot;
  // ajout mbo pour gestion prime ou subvention - pas de notion de base debut exercice de la cession
  if (MethAmort.PlanUtil = 'PRI') or (MethAmort.PlanUtil = 'SBV') then
  begin
     result := 0;
  end else
  begin
     if (ValeurAchat<>0) then Result := ValeurAchat - CumulDot
     else Result := BaseCalc - (CumulDot+MethAmort.RepriseCedee);    // au cas où valeur achat est nul (possible dans les anciennes version)
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : remarque concernant cette fonction
Créé le ...... : 22/09/2006
Modifié le ... :   /  /
Description .. : Cette fonction n'est utilisée qu'avec la méthode éco
Suite ........ : elle n'est donc pas modifiée pour la prime d'équipement
Mots clefs ... :
*****************************************************************}
function TPlanAmort.GetBaseExercice(DateFinExe: TDateTime; Methode:
  TMethodeAmort): double;
var
  Indice: integer;
begin
  Indice := 0;
  result := 0.00;
     while (TableauDate[Indice] <> DateFinExe) and
       (TableauDate[Indice] <> iDate1900) do
       Indice := Indice + 1;
     if (TableauDate[Indice] = DateFinExe) then
       result := Methode.BaseDebutExo[Indice];
end;

function TPlanAmort.GetBaseCessionExercice(DateFinExe: TDateTime; Methode:
  TMethodeAmort): double;
var
  Indice: integer;
begin
  Indice := 0;
  result := 0.00;
  while (TableauDate[Indice] <> DateFinExe) and
    (TableauDate[Indice] <> iDate1900) do
    Indice := Indice + 1;
  if (TableauDate[Indice] = DateFinExe) then
    result := Methode.BaseDebExoCes[Indice];
end;

procedure TPlanAmort.BasculeDotationCedee(PCede: TPlanAmort; DateOpe:
  TDateTime);
var
  i, j: integer;
  DateDebutEx, DateFinEx: TDateTime;
begin
  i := 0;
  while PCede.TableauDate[i] <> iDate1900 do
  begin
    j := 0;
    while (TableauDate[j] <> iDate1900) and (TableauDate[j] <>
      PCede.TableauDate[i]) do
      j := j + 1;
    if (TableauDate[j] <> iDate1900) then
    begin
      DateFinEx := iDate1900;
      GetDatesExercice(TableauDate[j], DateDebutEx, DateFinEx);
      begin
        {  AmortEco.BaseDebutExoCession[j]:=AmortEco.BaseDebutExoCession[j]+
                            PCede.AmortEco.BaseDebutExo[j]+PCede.AmortEco.TableauDot[i];}
          // Modif CA le 08/07/1999 - Calcul de la base debut exo cession = base de la partie cédée.
          // Ajout CA le 20/09/1999 - Prise en compte des sorties multiples.
        AmortEco.BaseDebExoCes[j] := AmortEco.BaseDebExoCes[j] +
          PCede.AmortEco.Base;
        // Ajout Montant Exc. de la partie cédée.
        AmortEco.TabCession[j] := AmortEco.TabCession[j] +
          PCede.AmortEco.TableauDot[i];
        AmortFisc.TabCession[j] := AmortFisc.TabCession[j] +
          PCede.AmortFisc.TableauDot[i];
      end;
    end;
    i := i + 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 04/09/2006
Modifié le ... : 04/09/2006
Description .. : equivalent de getCumulDot Exercice pour la
Suite ........ : méthode PRI (dotation prime d'équipement)
Suite ........ : Suivant la date envoyée on a soit le cumul des antérieurs
Suite ........ : soit antérieurs calculés + dotation de l'exercice
Suite......... : ATTENTION : appel avec bReprise = true : on prend en cpte les antérieurs saisis
Suite......... :             sinon on cumul uniquement les dotations antérieures
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.GetCumulPRI(DateRef: TDateTime; var CumulPRI
  : double; bCession, bReprise, bClot: boolean);
begin
  CumulPRI := AmortPri.GetCumulDotExercice((ImmoCloture and bClot), DateRef,
    TableauDate, bCession);
  if bReprise then  CumulPRI := CumulPRI + AmortPri.Reprise;

  if bCession then CumulPRI := CumulPRI + AmortPri.RepriseCedee;

end;

procedure TPlanAmort.GetCumulsDotExercice(DateRef: TDateTime; var CumulEco,
  CumulFisc: double; bCession, bReprise, bClot: boolean);
begin
  CumulEco := AmortEco.GetCumulDotExercice((ImmoCloture and bClot), DateRef,
    TableauDate, bCession);

  // impact antérieur dépréciation - mbo 02.06
  if bReprise then  CumulEco := CumulEco + AmortEco.Reprise + AmortEco.RepriseDep;


  CumulFisc := AmortFisc.GetCumulDotExercice((ImmoCloture and bClot), DateRef,
    TableauDate, bCession);

  // fq 15274 - pas d'impact depréciation sur fiscal en plan variable
  if bReprise then
  begin
    // impact antérieur dépréciation - mbo 02.06
    CumulFisc := CumulFisc + AmortFisc.Reprise;

    // fq 15274 - mbo - antérieurs sur immo en variable
    if (AmortEco.Methode <> 'VAR') then CumulFisc := CumulFisc + AmortFisc.RepriseDep;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : equivalent à GetCumulDotExercice pour dérogatoire et
Suite ........ : réintégration fiscale
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.GetCumulsDotExerciceDR(DateRef: TDateTime; var CumulDR,
  CumulFEC: double; bCession, bReprise, bClot: boolean);
begin
  CumulDR := AmortDerog.GetCumulDotExercice((ImmoCloture and bClot), DateRef,
    TableauDate, bCession);

  CumulFEC := AmortReint.GetCumulDotExercice((ImmoCloture and bClot), DateRef,
    TableauDate, bCession);

  if bReprise then
  begin
    CumulDR := CumulDR + AmortDerog.Reprise;
    CumulFEC := CumulFEC + AmortReint.Reprise;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 28/09/2006
Description .. : equivalent de getCumulDot Exercice pour la
Suite ........ : méthode SBV (dotation Subvention d'équipement)
Suite ........ : Suivant la date envoyée on a soit le cumul des antérieurs
Suite ........ : soit antérieurs calculés + dotation de l'exercice
Suite......... : ATTENTION : appel avec bReprise = true : on prend en cpte les antérieurs saisis
Suite......... :             sinon on cumul uniquement les dotations antérieures
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.GetCumulSBV(DateRef: TDateTime; var CumulSBV
  : double; bCession, bReprise, bClot: boolean);
begin
  CumulSBV := AmortSBV.GetCumulDotExercice((ImmoCloture and bClot), DateRef,
    TableauDate, bCession);
  if bReprise then
  begin
     if AmortSbv.ReprisePrem = false then
        CumulSBV := CumulSBV + AmortSBV.Reprise;
  end;

  // ajout pour correction FQ 19276
  if bCession then
     CumulSBV := CumulSBV + AmortSBV.RepriseCedee;
end;

//=======================================================

function TPlanAmort.GetDotationEncours(Methode: TMethodeAmort; bCession:
  boolean): double;
begin
  result := GetDotationExercice(DateFinEnCours, Methode, bCession);
end;

//========================================================

function TPlanAmort.GetDotationExercice(DateFinExe: TDateTime; Methode:
  TMethodeAmort; bCession: boolean): double;
var
  Indice: integer;
begin
  Indice := 0;
  result := 0.00;
  while (TableauDate[Indice] <> DateFinExe) and (TableauDate[Indice] <>
    iDate1900) do
    Indice := Indice + 1;
  if (TableauDate[Indice] = DateFinExe) then
    result := Methode.TableauDot[Indice];
  if bCession then
    result := result + Methode.TabCession[Indice];
end;

//=======================================================================

procedure TPlanAmort.SetMontantExc(Valeur: double);
begin
  MontantExc := Valeur;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 29/08/2002
Modifié le ... : 20/12/2004
Description .. : Calcul de la date théorique de fin d'amortissement
Suite ........ : - CA - 20/12/2004 - Prise en compte du cas de
Suite ........ : l'amortissement variable
Mots clefs ... :
*****************************************************************}

procedure TPlanAmort.CalculDateFinAmortissement(MethodeAmort: TMethodeAmort);
var
  DateDebutEx, DateFinEx: TDateTime;
  i, Duree, ResteMois: integer;
  PeriodePlus, bIncPeriode: boolean;
begin
  PeriodePlus := false;

  // Amortissement linéaire
  if MethodeAmort.Methode = 'LIN' then
  begin
    MethodeAmort.DateFinAmort := GetDateDebutAmort(MethodeAmort);
    Duree := MethodeAmort.Duree;
    i := 1;
    while (i <= Duree div 12) do
    begin
      MethodeAmort.DateFinAmort := PlusMois(MethodeAmort.DateFinAmort, 12);
      inc(i);
    end;
    ResteMois := Duree - ((i - 1) * 12);
    if (ResteMois > 0) then
      MethodeAmort.DateFinAmort := PlusMois(MethodeAmort.DateFinAmort,
        ResteMois);
    MethodeAmort.DateFinAmort := MethodeAmort.DateFinAmort - 1;
  end
    // Amortissement dégressif
  else if MethodeAmort.Methode = 'DEG' then
  begin
    // CA - 29/08/2002 - Si exercice de la date d'achat est inférieur à 12 mois,
    // on ajoute une période supplémentaire d'amortissement
    // mbo 27.04.06 - fq 17569  bIncPeriode := ExoAchatInferieurA12;
    if MethodeAmort.planUtil = 'FISC' then
       bIncPeriode := ExoAchatInferieurA12(DateDebFis)
    else if MethodeAmort.PlanUtil = 'ECO' then
       bIncPeriode := ExoAchatInferieurA12(DateDebEco)
    else if MethodeAmort.planUtil = 'PRI' then
       bIncPeriode := ExoAchatInferieurA12(PRIDateDeb)
    else
      bIncPeriode := ExoAchatInferieurA12(SBVDateDeb);

    DateDebutEx := iDate1900;
    DateFinEx := iDate1900;
    Duree := (MethodeAmort.Duree div 12);
    if Duree < (MethodeAmort.Duree / 12) then
      inc(duree)
    else if bIncPeriode then
      Inc(Duree);

    for i := 1 to Duree do
    { ajout mbo pour fq 15234 - 23/03/2006
        si l'exo d'achat est >= 12 mois
     et si un des exercices d'amortissement suivant l'exo d'achat a une durée inférieure à 12 mois
     alors il faut ajout une annuité supplémentaire (nvelle f° ExoInferieurA12)
     Le principe : on fait un premier tour à vide juste pour repérer si un exercice
     d'amortissement est inférieur à 12 mois pour pouvoir ajouter une annuité}
      begin
        GetDatesExercice(GetDateDebutAmort(MethodeAmort), DateDebutEx, DateFinEx);
        if bIncPeriode = false then
        begin
           bIncPeriode := ExoInferieurA12(DateDebutEx);   // nvelle f° qui renvoie true si exo <12 mois
           if bIncPeriode = true then
           begin
             PeriodePlus := true;
             AjoutPeriode:= true;  // fq 15234 - 15.06.06
           end;
        end;
      end;

    {on refait la boucle avec 1 période en plus si nécessaire }
    DateDebutEx := iDate1900;
    DateFinEx := iDate1900;

    if PeriodePlus = true then inc(Duree);

    for i := 1 to Duree do
    begin
        GetDatesExercice(GetDateDebutAmort(MethodeAmort), DateDebutEx, DateFinEx);
        // mbo - 06.06.2006 code inutile car bIncPeriode ne sera jamais égal à true
        //if bIncPeriode = false then
        //begin
        //   bIncPeriode := ExoInferieurA12(DateDebutEx);
        //   if bIncPeriode = true then inc(Duree);
        //end;
    end;

    MethodeAmort.DateFinAmort := DateFinEx;

  end else if MethodeAmort.Methode = 'VAR' then
  begin
    if (fPlanUO<>nil) and (fPlanUO.Detail.Count > 0) then
      MethodeAmort.DateFinAmort := fPlanUO.Detail[fPlanUO.Detail.Count-1].GetValue('IUO_DATE');
  end;
end;


procedure TPlanAmort.CalculDateFinAmortReprise(MethodeAmort: TMethodeAmort);
var
  i, Duree, ResteMois: integer;
begin
  MethodeAmort.DateFinAmort := GetDateDebutAmort(MethodeAmort);
  Duree := MethodeAmort.Duree;
  i := 1;
  while (i <= Duree div 12) do
  begin
    MethodeAmort.DateFinAmort := PlusMois(MethodeAmort.DateFinAmort, 12);
    inc(i);
  end;
  ResteMois := Duree - ((i - 1) * 12);
  if (ResteMois > 0) then
    MethodeAmort.DateFinAmort := PlusMois(MethodeAmort.DateFinAmort, 12 -
      ResteMois);
  MethodeAmort.DateFinAmort := MethodeAmort.DateFinAmort - 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 04/09/2006
Modifié le ... :   /  /
Description .. : NON UTILISEE
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.SetReprises(PlanOrig: TPlanAmort; Valeur: double);
begin
  AmortEco.Reprise := Valeur;
  AmortFisc.Reprise := Valeur;

end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 04/09/2006
Modifié le ... :   /  /
Description .. : NON UTILISEE
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.SetRevisionEco(Montant: double);
begin
  RevisionEco := Montant;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 04/09/2006
Modifié le ... :   /  /
Description .. : NON UTILISEE
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.SetRevisionFisc(Montant: double);
begin
  RevisionFisc := Montant;
end;

procedure TPlanAmort.RestaureRepriseSurImmoEclate(PEclate: TPlanAmort);
begin
  AmortEco.Reprise := AmortEco.Reprise + PEclate.AmortEco.Reprise;
  AmortFisc.Reprise := AmortFisc.Reprise + PEclate.AmortFisc.Reprise;

  AmortPri.Reprise := AmortPri.Reprise + PEclate.AmortPri.Reprise;
  AmortPri.RepriseDep := 0.00;

  AmortSbv.Reprise := AmortSbv.Reprise + PEclate.AmortSbv.Reprise;
  AmortSbv.RepriseDep := 0.00;

  // impact antérieur dépréciation - mbo 02.06
  AmortEco.RepriseDep := AmortEco.RepriseDep + PEclate.AmortEco.RepriseDep;
  AmortFisc.RepriseDep := AmortFisc.RepriseDep + PEclate.AmortFisc.RepriseDep;

end;


function TPlanAMort.GetDateFinAmort: TDateTime;
var DateFin:TDateTime;
begin
  if AmortEco.DateFinAmort > AmortFisc.DateFinAmort then
    DateFin := AmortEco.DateFinAmort
  else
    DateFin := AmortFisc.DateFinAmort;

  // ajout mbo suite présence prime ou/et subvention
  if AmortEco.Methode = 'NAM' then
  begin
     if AmortPri.Duree <> 0 then
     begin
       if (DateMiseEnService + (AmortPri.Duree * 30)) > DateFin then
            DateFin := (DateMiseEnService + (AmortPri.Duree * 30));
     end else
     begin
       if AmortSbv.Duree <> 0 then
         if (DateMiseEnService + (AmortSBV.Duree * 30)) > DateFin then
            DateFin := DateMiseEnService + (AmortSbv.Duree * 30);
     end;
  end;
  result := DateFin;
end;


function TPlanAmort.GetDotationGlobale(TypeMethode: integer; tDateCalcul:
  TDateTime; var dDot, dCes, dExc: double): double;
var
  Methode: TMethodeAmort;
  dDotTot: double;
  DateOpe: TDateTime;

begin

  if TypeMethode = MethodeEco then
    Methode := AmortEco
  else if TypeMethode = MethodeFisc then
    Methode := AmortFisc
  else if TypeMethode = MethodePri then
    Methode := AmortPri
  else if TypeMethode = MethodeSBV then
     Methode := AmortSbv
  else if TypeMethode = MethodeDerog then
     Methode := AmortDerog
  else
     Methode := AmortReint;

  dDot := GetDotationExercice(tDateCalcul, Methode, false);
  dCes := GetCessionExercice(tDateCalcul, Methode);

  // mbo 24.10.05 dExc := GetExcepExercice(tDateCalcul, DateOpe);
  // mbo 30/11/05 ajout du test pour ne pas prendre en compte la dépréciation dans excep fiscal
  // fq 17476 - la dépréciation impacte aussi le plan fiscal
  //if Methode = AmortEco then

  dExc := GetExcepExercice(tDateCalcul, DateOpe,false, true);
  //else
  //   dExc := GetExcepExercice(tDateCalcul, DateOpe,false, false);
  // modif mbo 27.10.05 on ne cumule plus revision qui = I_revisioneco = montant depreciation
  // dDotTot := dDot + dCes + Methode.Revision; // + dExc; (compris dans dDot) et depreciation comprise

  dDotTot := dDot + dCes;// + dExc; (compris dans dDot) et depreciation comprise

  // fq 17154 - ajout mbo 6.12.05
  if (methode = AmortFisc) and (EcoVar = true) then dExc := 0;

  if (methode = AmortPri) or (methode = AmortSbv) then dExc := 0;

  // ajout mbo 19.07.05 - FQ 15480 Ajout de l'amort exceptionnel car non pris en compte
  // ds le cas d'une immo dont la date de fin amort théorique est passée et qu'il reste à amortir
  { mbo - fq 18544 - 11/07/2006 si immo amortie ds l'exercice en cours : exceptionnel comptabilisé 2 fois
  if (((Methode.DateFinAmort >= DateDebutEncours) and (Methode.DateFinAmort <= DateFinEncours))  // Fin d'amortissement sur exercice en cours
      or  (Methode.DateFinAmort < DateDebutEncours)) // Fin d'amortissement thÚorique antÚrieur Ó la date de dÚbut d'exercice - FQ 15477 - CA - 16/03/2005
      then
        dDotTot := dDotTot + dExc;

  if (Methode.DateFinAmort < DateDebutEncours) // Fin d'amortissement thÚorique antÚrieur Ó la date de dÚbut d'exercice - FQ 15477 - CA - 16/03/2005
      then dDotTot := dDotTot + dExc;
  }
  result := dDotTot;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 25/08/2004
Modifié le ... :   /  /
Description .. : Récupère l'exceptionnel.
Suite ........ : bAvecSortie permet de garder la compatibilité avec
Suite ........ : l'antériorité ou le calcul de l'exceptionnel était fait à part (cas
Suite ........ : de la génération des écritures)
Mots clefs ... :
*****************************************************************}

function TPlanAmort.GetExcepExercice(tDateCalcul: TDateTime; var
  DateRevisionNam: TDateTime; bAvecSortie: boolean = false; bAvecDeprec: boolean = false): double;
var
  MontantDot, DotExcep: double;
  Q: TQuery;
  Requete, TypeDot: string;
  stWhereExc: string;
begin
  DotExcep := 0.00;
  result := 0.00;
  DateRevisionNam := iDate1900;
  if not ChangementPlan then
    exit; // Ajout CA le 28/07/1999 pour accélération des éditions
  if not ExerciceEnCours(tDateCalcul) then
    exit;
  if bAvecSortie then         // mbo 25.11.05 fq 16400 la cession n'était pas traitée
    begin
    stWhereExc := ' AND (IL_TYPEOP="ELE" OR IL_TYPEOP="CDM" OR ';
    stWhereExc := stWhereExc + 'IL_TYPEOP="CES" OR IL_TYPEOP="EEC" OR IL_TYPEOP="ELC" OR IL_TYPEEXC="DOT" OR IL_TYPEEXC="RDO" OR IL_TYPEOP="DPR")) '
    end
  else
    stWhereExc :=
      ' AND (IL_TYPEOP="ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="EEC" OR IL_TYPEOP="ELC" OR IL_TYPEOP="DPR")) ';
  Requete := 'SELECT IL_DATEOP,IL_MONTANTDOT,IL_TYPEDOT,IL_METHODEECO,IL_TYPEOP, IL_TYPEEXC, IL_MONTANTEXC FROM IMMOLOG '
    +
    'WHERE (IL_IMMO="' + CodeImmo + '"' +
    'AND IL_DATEOP>="' + USDateTime(DateDebutEnCours) + '"' +
    'AND IL_DATEOP<="' + USDatetime(tDateCalcul) + '"' +
    stWhereExc + ' ORDER BY IL_DATEOP ASC';
  Q := OpenSql(Requete, true);
  while not Q.EOF do
  begin
    { Si révision NAM vers amortissable ==> on récupère la date de début d'amortissement }
    if ((Q.FindField('IL_TYPEOP').AsString = 'CDM') and
      (Q.FindField('IL_METHODEECO').AsString = 'NAM')) then
      DateRevisionNam := Q.FindField('IL_DATEOP').AsDateTime;
    { Calcul du montant total de l'exceptionnel }
    if (Q.FindField('IL_TYPEOP').AsString = 'CES') then
    begin
      TypeDot := Q.FindField('IL_TYPEEXC').AsString;
      MontantDot := Q.FindField('IL_MONTANTEXC').AsFloat;
      if TypeDot = 'RDO' then     // ajout mbo 25.11.05
        DotExcep := DotExcep - MontantDot
      else
        DotExcep := DotExcep + MontantDot;
    end
    // TGA 24/11/05 ajout test EEC else if (Q.FindField('IL_TYPEOP').AsString = 'ELE') then
    else if ((Q.FindField('IL_TYPEOP').AsString = 'ELE') OR
             (Q.FindField('IL_TYPEOP').AsString = 'EEC')) then
    begin
      MontantDot := Q.FindField('IL_MONTANTDOT').AsFloat;
      TypeDot := Q.FindField('IL_TYPEDOT').AsString;
      if TypeDot = 'RDO' then
        DotExcep := DotExcep - MontantDot
          // CA - 25/08/2004 - Prise en compte du signe de la dotation
      else
        DotExcep := DotExcep + MontantDot;
    end;

    // MBO - 24.10.05 prise en compte de la depreciation d'actif
    if ((bAvecDeprec = true) and (Q.FindField('IL_TYPEOP').AsString = 'DPR')) then
       begin
       // depreciation := Q.FindField('IL_MONTANTDOT').AsFloat;
       DotExcep := DotExcep + Q.FindField('IL_MONTANTDOT').AsFloat;
    end;
    Q.Next;
  end;
  Ferme(Q);
  result := DotExcep;
end;


function TPlanAmort.TraiteMontantDotationProrata(DateOp, tDateCalcul: TDateTime;
  MontantDot: double): double;
var
  NbExe, NbDot: integer;
  PremMois, PremAnnee, NbMois, D, M, Y: Word;
begin
  DecodeDate(DateOp, Y, M, D);
  if (D = DaysPerMonth(Y, M)) then
    D := 30;
      // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  NOMBREMOIS(DateOp, DateFinEnCours, PremMois, PremAnnee, NbMois);
  NbExe := ((NbMois - 1) * 30 + 30 - MinIntValue([D, 30])) div 30;
  NOMBREMOIS(DateOp, tDateCalcul, PremMois, PremAnnee, NbMois);
  DecodeDate(tDateCalcul, Y, M, D);
  if (D = DaysPerMonth(Y, M)) then
    D := 30;
      // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  NbDot := ((NbMois - 1) * 30 + 30 - MinIntValue([D, 30])) div 30;
  result := (MontantDot) * (MaxIntValue([1, NbDot]) / MaxIntValue([1, NbExe]));
    //11/02
end;

//===============================================================================
// Récupération Elements exceptionnels Revision
function TPlanAmort.GetCessionExercice(tDateCalcul: TDateTime; Methode:
  TMethodeAmort): double;
var
  Indice: integer;
  DateDebutEx, DateFinEx: TDateTime;

begin
  Indice := 0;
  result := 0.00;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;

  GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);

  while (TableauDate[Indice] <> DateFinEx) and (TableauDate[Indice] <> iDate1900)
    do
    Indice := Indice + 1;
  if (TableauDate[Indice] = DateFinEx) and (Methode.TabCession[Indice] <> 0.00) then
  begin

    // fq 21405 - mbo 09.07 - ajout du test pour ne pas proratiser le dernier poste si cession après fin d'amortissement
    if DateDebutEnCours > tDateCalcul then
       result := 0
    else
    begin
       {//new 27.09
       if (Methode.planUtil = 'FISC') or (Methode.PlanUtil = 'ECO') then
       begin

          if (Methode.Methode='LIN') and (DateDebutEnCours = tDateCalcul) and (GetParamSocSecur('SO_AMORTJSORTIE', True) = 0) then
             result := 0
          else
             result := TraiteMontantDotationProrata(DateDebutEnCours, tDateCalcul,
                                                    Methode.TabCession[Indice]);
       end else
          result := TraiteMontantDotationProrata(DateDebutEnCours, tDateCalcul,
                                                    Methode.TabCession[Indice]);
       }


       result := TraiteMontantDotationProrata(DateDebutEnCours, tDateCalcul,Methode.TabCession[Indice]);
    end;
  end;
end;

//=====================================================================================

function TPlanAmort.ReintegrationFiscale: boolean;
begin
  result := (CoeffReintegration <> 0.00) or (TauxQuotePart <> 0.00);
end;

procedure TPlanAmort.GetElementExceptionnel(var Dot, DotExc: double; var
  sTypeExc: string);
begin
  //  if TypeExc='SDO' then Dot := 0.00
  if TypeExc = 'RDO' then
    DotExc := -MontantExc
  else
    DotExc := MontantExc;
  //  DotExc:=MontantExc;
  sTypeExc := TypeExc;
end;

procedure TPlanAmort.AffecteDotationCedee(PCedee: TPlanAMort; tDateCalcul:
  TDateTime);
var
  Indice: integer;
  DateDebutEx, DateFinEx: TDateTime;
begin
  Indice := 0;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  AmortEco.Base := PCedee.AmortEco.Base;
  AmortFisc.Base := PCedee.AmortFisc.Base;
  AmortPri.Base:= PCedee.AmortPri.Base;
  AmortSbv.Base:= PCedee.AmortSbv.Base;

  GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);

  while (PCedee.TableauDate[Indice] <= DateFinEx) and
    //EPZ 24/11/00 pb ecriture non générées....!       (PCedee.AmortEco.TableauDot[Indice]<>0.00) and
  (PCedee.TableauDate[Indice] <> iDate1900) do
  begin
    AmortEco.BaseDebExoCes[Indice] := PCedee.AmortEco.Base;
    if PCedee.TableauDate[Indice] = DateFinEx then
    begin
      AmortEco.TabCession[Indice] := AmortEco.TabCession[Indice] +
        PCedee.AmortEco.TableauDot[Indice];
      AmortFisc.TabCession[Indice] := AmortFisc.TabCession[Indice] +
        PCedee.AmortFisc.TableauDot[Indice];

      AmortPri.TabCession[Indice] := AmortPri.TabCession[Indice] +
        PCedee.AmortPri.TableauDot[Indice];
      AmortSbv.TabCession[Indice] := AmortSbv.TabCession[Indice] +
        PCedee.AmortSbv.TableauDot[Indice];

    end
    else
    begin
      AmortEco.TabCession[Indice] := PCedee.AmortEco.TableauDot[Indice];
      AmortFisc.TabCession[Indice] := PCedee.AmortFisc.TableauDot[Indice];

      AmortPri.TabCession[Indice] := PCedee.AmortPri.TableauDot[Indice];
      AmortSbv.TabCession[Indice] := PCedee.AmortSbv.TableauDot[Indice];

    end;
    Indice := Indice + 1;
  end;
end;

//Principe : calculer le prorata de la dotation (date de sortie)
// et réinitialiser le reste du tableau de dotations

procedure TPlanAmort.CalculProrataDotationCedee(MethodeAmort: TMethodeAmort;
  tDateCalcul: TDateTime);
var
  Indice: integer;
  dProrata, dMontantDot: double;
  DateDebutEx, DateFinEx: TDateTime;
  Year, Month, Day, DayFin: Word;
  dtFinAmort, dtDebutCalc: TDateTime;

  dMontantExc: double;
  sTypeExc: string;

begin
  Indice := 0;
  // #### Debut Ajout CA le 01/08/2000
  // Pour résoudre le problème de Dotation mal calculée pour cession après date
  // de fin d'amortissement. La Valeur de IA_CESSIONECO était incorrecte.
  bDotationExercice := False;
  // ### Fin Ajout CA le 01/08/2000
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  // CA le 06/09/1999 - calcul prorata par rapport à la date de fin d'amortissement
  // si fin amortissement durant l'exercice.
  dtFinAmort := AmortEco.DateFinAmort;
  if (dtFinAmort < DateDebutEnCours) or (dtFinAmort > DateFinEnCours) then
    dtFinAmort := DateFinEnCours
  else
  begin
    if tDateCalcul > dtFinAmort then
      tDateCalcul := dtFinAmort;
  end;
  GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
  while (TableauDate[Indice] <> iDate1900) and
    (TableauDate[Indice] <= DateFinEx) and
    (MethodeAmort.TableauDot[Indice] <> 0.00) do
  begin
    // On retire le montant des exceptionnels
    dMontantDot := MethodeAmort.TableauDot[Indice];

    if (MethodeAmort.PlanUtil = 'ECO') or (MethodeAmort.PlanUtil = 'FISC') then
    begin
      GetElementExceptionnel(dMontantDot, dMontantExc, sTypeExc);
      dMontantDot := dMontantDot - dMontantExc;
    end;

    if (TableauDate[indice] >= DateDebutEx) then
      // On proratise uniquement pour l'exercice en cours.
    begin
      (*if (TypeAmortCession<>'NOR') then //22/03/99 ne pas faire si amort total*)
      begin
        // CA - 02/07/99 - Si dernier jour du mois, on calcul sur le mois complet
        DecodeDate(tDateCalcul, Year, Month, Day);
        DecodeDate(FinDeMois(tDateCalcul), Year, Month, DayFin);
        if Day = DayFin then
          Day := 0;
        if (GetDateDebutAmort(MethodeAmort) > DateDebutEnCours) and
          (GetDateDebutAmort(MethodeAmort) < DateFinEnCours) then
          dtDebutCalc := GetDateDebutAmort(MethodeAmort)
        else
          dtDebutCalc := DateDebutEnCours;
        if (MethodeAmort.Methode = 'LIN') or (MethodeAmort.Methode='VAR')then
          dMontantDot := GetDotationLinAvant(dtDebutCalc, dtFinAmort,
            GetDateDebutAmort(MethodeAmort), tDateCalcul,
            MethodeAmort.DateFinAmort, dMontantDot, bDotationExercice)
        else if MethodeAmort.Methode = 'DEG' then
          dMontantDot := GetDotationDegAvant(dtDebutCalc, dtFinAmort,
            GetDateDebutAmort(MethodeAmort), tDateCalcul - Day, dMontantDot,
            dProrata, SurAmort);
      end;
    end;
    MethodeAmort.TableauDot[Indice] := dMontantDot;
    //FIN 02/04/99
    Indice := Indice + 1;
  end;
  //  Indice:=Indice+1;
  while MethodeAmort.TableauDot[Indice] <> 0.00 do
  begin
    TableauDate[Indice] := iDate1900;
    MethodeAmort.TableauDot[Indice] := 0.00;
    Indice := Indice + 1;
  end;
end;

procedure TPlanAmort.CalculProrataDotationOrigine(MethodeAmort: TMethodeAmort;
  Prorata: double);
var
  Indice: integer;
begin
  Indice := 0;
  while (TableauDate[Indice] <> iDate1900) and
    (TableauDate[Indice] < DateDebutEnCours) and
    (MethodeAmort.TableauDot[Indice] <> 0.00) do
  begin
    MethodeAmort.TableauDot[Indice] := MethodeAmort.TableauDot[Indice] *
      Prorata;
    Inc(Indice);
  end;
end;

// Procedure de chargement de plan à partir d'un objet TOB

procedure TPlanAmort.ChargePlanTOB(OB: TOB);
Var
  base : double;
  bfiscal : boolean;
begin
  NumSeq := OB.GetValue('I_PLANACTIF');
  CodeImmo := OB.GetValue('I_IMMO');
  LibelleImmo := OB.GetValue('I_LIBELLE');
  CompteImmo := OB.GetValue('I_COMPTEIMMO');
  NbMoisReprise := OB.GetValue('I_DUREEREPRISE');
  DateAchat := OB.GetValue('I_DATEPIECEA');
  DateMiseEnService := OB.GetValue('I_DATEAMORT');
  BaseDebutExoEco := OB.GetValue('I_BASEECO');
  BaseDebutExoFisc := OB.GetValue('I_BASEFISC');
  BaseDebutAmort := OB.GetValue('I_BASEAMORDEBEXO');
  BaseFinAmort := OB.GetValue('I_BASEAMORFINEXO');
  DateCreation := OB.GetValue('I_DATECREATION');
  MontantExc := OB.GetValue('I_MONTANTEXC');
  TypeExc := OB.GetValue('I_TYPEEXC');
  ValReintegration := OB.GetValue('I_REINTEGRATION');
  TauxQuotePart := OB.GetValue('I_QUOTEPART');
  ValeurHT := OB.GetValue('I_MONTANTHT');
  ValeurTvaRecuperable := OB.GetValue('I_TVARECUPERABLE');
  ValeurTvaRecuperee := OB.GetValue('I_TVARECUPEREE');

  ValeurAchat := OB.GetValue('I_VALEURACHAT');
 // MVG 08/2006 DPI
  DPI := (OB.GetValue('I_DPI') = 'X');
  DPIAffectee := MontantDPIAffectee (CodeImmo);

  // Ajout mbo 10.09.07 - FQ 21405
  CedeLe := OB.GetValue('I_DATECESSION');  // I_DATECREATION'); // = I_DATECESSION

  // MVG 30/08/2006 - FQ 18720 : le calcul se fait par rapport à la base éoonomique
  qte := OB.GetValue('I_QUANTITE');

  Base :=Recup_Baseeco(Codeimmo,qte = 0.0);
  if OB.GetValue('I_METHODEFISC')<>'' then bfiscal:=true else bfiscal:=false;
  if bfiscal then Base :=Recup_Basefisc(Codeimmo,qte = 0.0);
  // MVG 30/08/2006 - FQ 18720 : le calcul se fait par rapport à la base éoonomique
  if base <> 0.00 then
    CoeffReintegration := ((OB.GetValue('I_REINTEGRATION') + DPIAffectee) /
      Base);
  ImmoCloture := OB.GetValue('I_ETAT') <> 'OUV'; //EPZ 29/03/99
  ChangementPlan := OB.GetValue('I_OPECHANGEPLAN') = 'X';
  Suramort := OB.GetValue('I_SURAMORT') = 'X';

  { Chargement du plan d'unités d'oeuvre }
  fPlanUO.LoadDetailFromSQL('SELECT * FROM IMMOUO WHERE IUO_IMMO="' + CodeImmo +
    '"');
  fTotalUO := OB.GetValue('I_UNITEOEUVRE');

  // mbo 27.10.05
  RevisionEco := OB.GetValue('I_REVISIONECO');
  CumulDepre := Cumul_depreciation(Codeimmo,false);
  journala  := OB.GetValue('I_JOURNALA') ;
  IF OB.GetValue('I_JOURNALA') = '***' then
     VncRestEco := True
  Else
     VncRestEco := False;

  // ajout pour chantier fiscal
  VncRestFisc := OB.GetValue('I_FUTURVNFISC') = '***';

  IF OB.GetValue('I_REVISIONECO') <> 0 then
   DureeRest := True
  ELSE
   DureeRest := False;

  // fa 17569 ajout mbo 27.04.06
  DateDebEco := OB.GetValue('I_DATEDEBECO');
  DateDebFis := OB.GetValue('I_DATEDEBFIS');

  // Ajout mbo pour gestion prime d'équipement
  PRI := (OB.GetValue('I_SBVPRI') <> 0) or (OB.GetValue('I_SBVPRIC') <> 0) ;
  if OB.GetValue('I_SBVPRI') <> 0 then
    MNTPrime:=OB.GetValue('I_SBVPRI')
  else
    MNTPrime:=OB.GetValue('I_SBVPRIC');

  if (PRI) then
  begin
     if (bFiscal) then
     begin
        PRIDuree:=OB.GetValue('I_DUREEFISC');
        PRITaux:=OB.GetValue('I_TAUXFISC');
        PRIMethode:=OB.GetValue('I_METHODEFISC');
        PRIDateDeb := DateDebFis;
     end else
     begin
        if (OB.GetValue('I_METHODEECO') <> 'NAM') then
        begin
           PRIDuree:=OB.GetValue('I_DUREEECO');
           PRITaux:=OB.GetValue('I_TAUXECO');
           PRIMethode:=OB.GetValue('I_METHODEECO');
           PRIDateDeb := DateDebEco;
        end else
        begin
           PRIDuree:=OB.GetValue('I_CORVRCEDDE');
           PRITaux:= GetTaux('LIN',DateMiseEnService, DateMiseEnService, PRIDuree);
           PRIMethode:='LIN';
           PRIDateDeb := StrToDate(OB.GetValue('I_DATEAMORT'));
        end;
     end;
  end;

  SBV := (OB.GetValue('I_SBVMT') <> 0) or (OB.GetValue('I_SBVMTC') <> 0);

 if (SBV) then
 begin
  // correction mbo FQ 19276
  //MntSBV:=OB.GetValue('I_SBVMT');
  if OB.GetValue('I_SBVMT') <> 0 then
    MntSBV:=OB.GetValue('I_SBVMT')
  else
    MntSbv:=OB.GetValue('I_SBVMTC');

  RepriseCedeeSBV := Recup_RepriseCedeeSBV(Codeimmo,qte = 0.0);

     SBVDateSBV := OB.GetValue('I_SBVDATE');
     if (bFiscal) then
     begin
        SBVDuree:=OB.GetValue('I_DUREEFISC');
        SBVTaux:=OB.GetValue('I_TAUXFISC');
        SBVMethode:=OB.GetValue('I_METHODEFISC');
        SBVDateDeb:= DateDebFis;
     end else
     begin
        if (OB.GetValue('I_METHODEECO') <> 'NAM') then
        begin
           SBVDuree:=OB.GetValue('I_DUREEECO');
           SBVTaux:=OB.GetValue('I_TAUXECO');
           SBVMethode:=OB.GetValue('I_METHODEECO');
           SBVDateDeb:= DateDebEco;
        end else
        begin
           SBVDuree:=OB.GetValue('I_CORVRCEDDE');
           SBVTaux:= GetTaux('LIN',DateMiseEnService, DateMiseEnService, SBVDuree);
           SBVMethode:='LIN';
           SBVDateDeb:= DateMiseEnService;
        end;
     end;
 end;

 // ajout chantier fiscale
  GestionFiscale := (OB.GetValue('I_NONDED') = 'X');
  RepriseDerog := OB.GetValue('I_REPRISEDR');
  RepriseReint := OB.GetValue('I_REPRISEFEC');
  RepriseDRCedee := OB.GetValue('I_REPRISEFDRCEDEE');
  RepriseRICedee := OB.GetValue('I_REPRISEFECCEDEE');
  EstRemplacee := (OB.GetValue('I_REMPLACEE')<> '');
  Remplace := (OB.GetValue('I_REMPLACE')<>'');

end;

procedure TPlanAmort.ChargeMethodeTOB(OB: TOB);
var
  Clot: boolean;
  dTaux: double;
  dDuree: integer;
  ReprisePrem: boolean;
begin
  Clot := OB.GetValue('I_ETAT') <> 'OUV';

  // ajout mbo 09.10.06 pour reprise subvention sur la première dotation calculée
  if OB.GetValue('I_DPIEC') = '-' then
     ReprisePrem := false
  else
     ReprisePRem := true;

  { Recalcul du taux en fonction de la durée d'amortissement }
  if ((OB.GetValue('I_DUREEECO') <> 0) or (OB.GetValue('I_TAUXECO') = 0)) then
  begin
    { cas ou l'on connait la durée, on veut recalculer le taux
     FQ 17569 mbo modif date debut amort
    dTaux := GetTaux(OB.GetValue('I_METHODEECO'), OB.GetValue('I_DATEAMORT'),
      OB.GetValue('I_DATEPIECEA'), OB.GetValue('I_DUREEECO'));
    }
    dTaux := GetTaux(OB.GetValue('I_METHODEECO'), OB.GetValue('I_DATEDEBECO'),
             OB.GetValue('I_DATEDEBECO'), OB.GetValue('I_DUREEECO'));
    OB.PutValue('I_TAUXECO', dTaux);
  end
  else
  begin
    { cas où l'on connait le taux et on veut recalculer la durée
    FQ 17569
    dDuree := AMTaux2Duree(OB.GetValue('I_METHODEECO'),
              OB.GetValue('I_DATEPIECEA'), OB.GetValue('I_TAUXECO'));
    }
    dDuree := AMTaux2Duree(OB.GetValue('I_METHODEECO'),
              OB.GetValue('I_DATEDEBECO'), OB.GetValue('I_TAUXECO'));
    OB.PutValue('I_DUREEECO', dDuree);
  end;

  AmortEco.AffecteMethodeAmort(TableauDate, OB.GetValue('I_DATECESSION'),
    OB.GetValue('I_REVISIONECO'), 0.00 (*Exceptionnel*),
    OB.GetValue('I_REPRISEECO'),
    OB.GetValue('I_REPCEDECO'), OB.GetValue('I_BASEECO'),
    OB.GetValue('I_TAUXECO'),
    OB.GetValue('I_DUREEECO'), OB.GetValue('I_METHODEECO'), Clot, 'ECO',
    OB.GetValue('I_REPRISEDEP'), false, false, 0);

  Fiscal := OB.GetValue('I_METHODEFISC') <> '';

  if Fiscal then //EPZ 30/11/00
  begin
    { FQ 17569 mbo
    dTaux := GetTaux(OB.GetValue('I_METHODEFISC'), OB.GetValue('I_DATEAMORT'),
    OB.GetValue('I_DATEPIECEA'), OB.GetValue('I_DUREEFISC'));
    }
    dTaux := GetTaux(OB.GetValue('I_METHODEFISC'), OB.GetValue('I_DATEDEBFIS'),
             OB.GetValue('I_DATEDEBFIS'), OB.GetValue('I_DUREEFISC'));
    OB.PutValue('I_TAUXFISC', dTaux);

    AmortFisc.AffecteMethodeAmort(TableauDate, OB.GetValue('I_DATECESSION'),
      OB.GetValue('I_REVISIONFISCALE'), 0.00 (*Exceptionnel*),
      OB.GetValue('I_REPRISEFISCAL'),
      OB.GetValue('I_REPCEDFISC'), OB.GetValue('I_BASEFISC'),
      OB.GetValue('I_TAUXFISC'),
      OB.GetValue('I_DUREEFISC'), OB.GetValue('I_METHODEFISC'), Clot, 'FISC',
      OB.GetValue('I_REPRISEDEP'), false, false, 0);

    // ajout pour chantier fiscal
    AmortDerog.AffecteMethodeAmort(TableauDate, OB.GetValue('I_DATECESSION'),
      0.00 {*'I_REVISIONFISCALE'*}, 0.00 (*Exceptionnel*),
      OB.GetValue('I_REPRISEDR'),
      OB.GetValue('I_REPRISEFDRCEDEE'), 0.00  {*I_BASEFISC*},
      0.00 {*TAUX*},
      0 {*DUREE*}, '' {*METHODE*}, Clot, 'DRG',
      0.00 {*REPRISEDEP*}, false, false, 0);

    AmortReint.AffecteMethodeAmort(TableauDate, OB.GetValue('I_DATECESSION'),
      0.00 {*'I_REVISIONFISCALE'*}, 0.00 (*Exceptionnel*),
      OB.GetValue('I_REPRISEFEC'),
      OB.GetValue('I_REPRISEFECCEDEE'), 0.00  {*I_BASEFISC*},
      0.00 {*TAUX*},
      0 {*DUREE*}, '' {*METHODE*}, Clot, 'REI',
      0.00 {*REPRISEDEP*}, false, false, 0);
  end;

   PRI := (OB.Getvalue('I_SBVPRI') <> 0.00) or (OB.Getvalue('I_SBVPRIC') <> 0.00);
   if (PRI) then
   begin
     if (fiscal) then
     begin
        AmortPri.AffecteMethodeAmort(TableauDate,
           OB.GetValue('I_DATECESSION'),
           0.00 (*cumul dépréciation*),
           0.00 (*Exceptionnel*),
           OB.GetValue('I_REPRISEUO'),
           Arrondi((OB.GetValue('I_REPRISEUOCEDEE')),V_PGI.OkDecV),
           Arrondi((OB.GetValue('I_SBVPRI')/2),V_PGI.OkDecV),
           OB.GetValue('I_TAUXFISC'), OB.GetValue('I_DUREEFISC'),
           OB.GetValue('I_METHODEFISC'), Clot, 'PRI',0, false, false, 0);  // on met reprisedep à 0
     end else
     begin
        if (OB.GetValue('I_METHODEECO') = 'NAM') then
        begin
           AmortPri.AffecteMethodeAmort(TableauDate,
           OB.GetValue('I_DATECESSION'),
           0.00 (*cumul dépréciation*),
           0.00 (*Exceptionnel*),
           OB.GetValue('I_REPRISEUO'),
           Arrondi((OB.GetValue('I_REPRISEUOCEDEE')),V_PGI.OkDecV),
           Arrondi((OB.GetValue('I_SBVPRI')/2),V_PGI.OkDecV),
           GetTaux('LIN', OB.GetValue('I_DATEAMORT'), OB.GetValue('I_DATEAMORT'),
                    Trunc(OB.GetValue('I_CORVRCEDDE'))),
           Trunc(OB.GetValue('I_CORVRCEDDE')),
           'LIN', Clot, 'PRI', 0.00, false, false, 0) (*reprise depreciation*);
        end else
        begin
           AmortPri.AffecteMethodeAmort(TableauDate,
           OB.GetValue('I_DATECESSION'),
           0.00 (*cumul dépréciation*),
           0.00 (*Exceptionnel*),
           OB.GetValue('I_REPRISEUO'),
           Arrondi((OB.GetValue('I_REPRISEUOCEDEE')),V_PGI.OkDecV),
           Arrondi((OB.GetValue('I_SBVPRI')/2),V_PGI.OkDecV),
           OB.GetValue('I_TAUXECO'), OB.GetValue('I_DUREEECO'),
           OB.GetValue('I_METHODEECO'), Clot, 'PRI',0, false, false, 0);  // on met reprisedep à 0
        end;
     end;
   end;

   SBV := (OB.Getvalue('I_SBVMT') <> 0.00) or  (OB.Getvalue('I_SBVMTC') <> 0.00) ;
   if (SBV) then
   begin
     if (fiscal) then
     begin
        AmortSBV.AffecteMethodeAmort(TableauDate,
           OB.GetValue('I_DATECESSION'),
           0.00 (*cumul dépréciation*),
           0.00 (*Exceptionnel*),
           OB.GetValue('I_CORRECTIONVR'),
           RepriseCedeeSBV,  // reprise anterieure cedee
           OB.GetValue('I_SBVMT'),
           OB.GetValue('I_TAUXFISC'), OB.GetValue('I_DUREEFISC'),
           OB.GetValue('I_METHODEFISC'), Clot, 'SBV',0, false, ReprisePrem, 2);  // on met reprisedep à 0
     end else
     begin
        if (OB.GetValue('I_METHODEECO') = 'NAM') then
        begin
           AmortSBV.AffecteMethodeAmort(TableauDate,
           OB.GetValue('I_DATECESSION'),
           0.00 (*cumul dépréciation*),
           0.00 (*Exceptionnel*),
           OB.GetValue('I_CORRECTIONVR'),
           RepriseCedeeSBV,
           OB.GetValue('I_SBVMT'),
           GetTaux('LIN', OB.GetValue('I_DATEAMORT'), OB.GetValue('I_DATEAMORT'),
                                Trunc(OB.GetValue('I_CORVRCEDDE'))),
           Trunc(OB.GetValue('I_CORVRCEDDE')),
           'LIN', Clot, 'SBV', 0.00, false, ReprisePrem, 0) (*reprise depreciation*);
        end else
        begin
           AmortSBV.AffecteMethodeAmort(TableauDate,
           OB.GetValue('I_DATECESSION'),
           0.00 (*cumul dépréciation*),
           0.00 (*Exceptionnel*),
           OB.GetValue('I_CORRECTIONVR'),
           RepriseCedeeSBV,
           OB.GetValue('I_SBVMT'),
           OB.GetValue('I_TAUXECO'), OB.GetValue('I_DUREEECO'),
           OB.GetValue('I_METHODEECO'), Clot, 'SBV',0, false, ReprisePrem, 1);  // on met reprisedep à 0
        end;
     end;
   end;
end;

procedure TPlanAmort.ChargeTOB(OB: TOB);
begin
  ChargePlanTOB(OB);
  ChargeMethodeTOB(OB);
end;

procedure TPlanAmort.CalculTOB(OB: TOB; tDateCalcul: TDateTime; bRecalculReprise : boolean = False);
var
  DateDebutExoRef: TDateTime;
  wMeth: string;
begin
  bDotationExercice := true;
  if tDateCalcul = iDate1900 then
    ChargeTOB(OB);
  { FQ 15673 - Raz des des reprises le cas échéant ( changement de date de début d'exercice uniquement ) }
  if bRecalculReprise then
  begin
    CalculReprises;
    OB.PutValue('I_REPRISEECO',AmortEco.Reprise);
    OB.PutValue('I_REPRISEFISCAL',AmortFisc.Reprise);

    OB.PutValue('I_REPRISEDR',AmortDerog.Reprise);   // ajout mbo fq 17512
    OB.PutValue('I_REPRISEFEC',AmortReint.Reprise);  // ajout mbo fq 17512

    OB.PutValue('I_DUREEREPRISE',CalculDureeReprise('ECO'));

    OB.PutValue('I_REPRISEUO',AmortPri.Reprise);
    OB.PutValue('I_CORRECTIONVR', AmortSbv.Reprise);

  end;

  DateCalcul := tDateCalcul;
  //FQ 17569 DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEPIECEA'));
  // fq 18333
  // MVG  FQ 18471 22/06/2006 if OB.GetValue('I_METHODEFISC').AsString <> ''  then vu C.AYEL
  if OB.GetString('I_METHODEFISC') <> ''  then
  begin
     if OB.GetValue('I_DATEDEBECO') > OB.GetValue('I_DATEDEBFIS') then
        DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEDEBFIS'))
     else
        DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEDEBECO'));
  end else
  begin
     DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEDEBECO'));
  end;

  ResetTableauDot(AmortEco, DateDebutExoRef);
  CalculExeptionnel(AmortEco, DateDebutExoRef);
  wMeth := OB.GetValue('I_METHODEECO');

  { fq 17569 mbo
  if wMeth = 'LIN' then
    DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEAMORT'))
  else
    DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEPIECEA'));
  }

  DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEDEBECO'));
  CalculDotation(AmortEco, DateDebutExoRef, false);

  if Fiscal then
  begin
    wMeth := OB.GetValue('I_METHODEFISC');

    { fq 17569 mbo
    if wMeth = 'LIN' then
      DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEAMORT'))
    else
      DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEPIECEA'));
    }
    DateDebutExoRef := GetDateDebutExoRef(OB.GetValue('I_DATEDEBFIS'));
    ResetTableauDot(AmortFisc, DateDebutExoRef);
    // fq 17102 - ajout mbo 29/11/05
    if EcoVar = false then   // ajout mbo 16.02.06
       CalculExeptionnel(AmortFisc, DateDebutExoRef);
    CalculDotation(AmortFisc, DateDebutExoRef, false);

    // ajout chantier fiscal
    ResetTableauDot(AmortDerog, DateDebutExoRef);
    ResetTableauDot(AmortReint, DateDebutExoRef);
    Calcul_Derog_Reint;
  end;

  if PRI then
  begin
    wMeth := AmortPri.Methode;
    DateDebutExoRef := GetDateDebutExoRef(PRIDateDeb);
    ResetTableauDot(AmortPri, DateDebutExoRef);
    CalculDotation(AmortPri, DateDebutExoRef, false);
  end;
  if SBV then
  begin
    wMeth := AmortSbv.Methode;
    DateDebutExoRef := GetDateDebutExoRef(SBVDateDeb);
    ResetTableauDot(AmortSbv, DateDebutExoRef);
    CalculDotation(AmortSbv, DateDebutExoRef, false);
  end;

end;

procedure TPlanAmort.SauveTOB(OBPlan: TOB);
var
  DatePrec, DateFinAmort: TDateTime;
  Methode: TMethodeAmort;
  Lig: integer;
  CessEco, CessFisc, CumulEco, CumulFisc, DotDerog, DotEco, DotFisc: double;
  DotPri, CessPri : double;
  DotSbv, CessSbv : double;
  OB: TOB;
  D2, D3, D4 : double;

  DateDebut, DateFin : TDateTime;

  // ajout pour chantier fiscal
  DotReint : double;
  CessionNd : double;
begin
  DatePrec := iDate1900;
  DateDebut := iDate1900;
  DateFin := iDate1900;
  NumSeq := NumSeq + 1;
  Modifie := false;

  { FQ 13259 - CA - 18/03/2005 - Prise en compte de la reprise cédée pour les immos cédées }
  CumulEco := AmortEco.Reprise+AmortEco.RepriseCedee;
  CumulFisc := AmortFisc.Reprise+AmortFisc.RepriseCedee;

  DateFinAmort := GetDateFinAmort;

  // ajout mbo pour éviter les lignes à 0 en fin d'immoamor
  // recherche fin exercice de l'exercice de fin d'amortissement
  GetDatesExercice(DateFinAmort, DateDebut, DateFin);
  for Lig := 0 to MAX_LIGNEDOT do
  begin
    // FQ 19530 - MBO - 16.01.2007
    D2 := 0;
    D3 := 0;
    D4 := 0;

    if TableauDate[Lig] > DateFin then
    begin
       If AmortFisc.Methode <> '' then D2 := AmortFisc.TableauDot[Lig];
       if AmortPRI.Methode <> '' then D3:= AmortPRI.TableauDot[Lig];
       if AmortSBV.Methode <> '' then D4:= AmortSBV.TableauDot[Lig];

       if (AmortEco.TableauDot[lig]=0) and (D2=0) and (D3=0) and (D4=0) then
          TableauDate[Lig] := iDate1900;
    end;
  end;

  Lig := 0;


while TableauDate[Lig] <> iDate1900 do
begin
  OB := TOB.Create('IMMOAMOR', OBPlan, -1);
  OB.PutValue('IA_IMMO', CodeImmo);
  OB.PutValue('IA_COMPTEIMMO', CompteImmo);
  OB.PutValue('IA_NUMEROSEQ', NumSeq);
  OB.PutValue('IA_CHANGEAMOR', TypeOpe);
  //     OB.PutValue('IA_BASEDEBEXOECO').AsFloat:=BaseDebutExoEco;
  //     OB.PutValue('IA_BASEDEBEXOFISC').AsFloat:=BaseDebutExoFisc;
  //     OB.PutValue('IA_DATEFIN').AsDateTime:=DateFinAmort;  // Modif CA le 09/03/1999
  //     OB.PutValue('IA_DATEFIN',DateCalcul);
  if TableauDate[Lig] <> iDate1900 then
  begin
    OB.PutValue('IA_DATE', TableauDate[Lig]);
    {Dotation économique}
    Methode := AmortEco;
    DotEco := Methode.TableauDot[Lig];
    OB.PutValue('IA_MONTANTECO', DotEco);
    OB.PutValue('IA_REINTEGRATION', DotEco * CoeffReintegration);
    OB.PutValue('IA_QUOTEPART', DotEco * TauxQuotePart);

    {Dotation partie(s) cédée(s)}
    CessEco := 0.00;
    if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec)
      then CessEco := Methode.TabCession[Lig];

    DotFisc := 0.00;
    DotDerog := 0.00;
    DotReint := 0.00;
    CessionNd := 0.00;
    CessFisc := 0.00;

    if AmortFisc.Methode <> '' then
    begin
      Methode := AmortFisc;
      DotFisc := Methode.TableauDot[Lig];
      if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec)
         then CessFisc := Methode.TabCession[Lig];

      DotDerog := AmortDerog.TableauDot[Lig];
      DotReint := AmortReint.TableauDot[Lig];
      if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec)
         then CessionNd := AmortReint.TabCession[Lig];

      end;
      {Dotation fiscale}
      OB.PutValue('IA_MONTANTFISCAL', DotFisc);
      {Dotation Dérogatoire}
      OB.PutValue('IA_MONTANTDEROG', DotDerog);
      // ajout pour chantier fiscal    réintégration extra comptable
      OB.PutValue('IA_NONDEDUCT', DotReint);
      OB.PutValue('IA_CESSIONND', CessionNd);

      {Dotation partie(s) cédée(s)}
      OB.PutValue('IA_CESSIONECO', CessEco);
      { CA - 20/01/2004 - FQ 13217 - Calcul de la réintégratin fiscale en cas de sortie }
      if (ValeurAchat <> 0) then
        OB.PutValue('IA_REINTEGRATION', CessEco * (ValReintegration /
          ValeurAchat));

      OB.PutValue('IA_CESSIONFISCAL', CessFisc);

      OB.PutValue('IA_BASEDEBEXOECO', AmortEco.Base - CumulEco);
      OB.PutValue('IA_BASEDEBEXOFISC', AmortFisc.Base - CumulFisc);
      //23/04/99
      //          CumulEco := CumulEco + DotEco + CessEco;
      //          CumulFisc := CumulFisc + DotFisc + CessFisc;
      CumulEco := CumulEco + DotEco;
      CumulFisc := CumulFisc + DotFisc;
      OB.PutValue('IA_BASEDEBEXOCESS', AmortEco.BaseDebExoCes[Lig]);
      { FQ 13478 - 17/03/2005 - CA - pour éviter de perdre la sortie au deuxième recalcul !!!}
      if (AmortEco.BaseDebExoCes[Lig]>0) then OB.PutValue('IA_BASEDEBEXOECO',AmortEco.BaseDebExoCes[Lig]);

    if AmortPri.Methode <> '' then
    begin
       Methode := AmortPri;
       DotPri := Methode.TableauDot[Lig];
       if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
          CessPri := Methode.TabCession[Lig]
       else
          CessPri := 0;

       OB.PutValue('IA_MONTANTSBV', DotPri);  {Dotation pri}
       OB.PutValue('IA_CESSIONSBV', CessPri); {Cession pri}
    end;

    if AmortSbv.Methode <> '' then
    begin
       Methode := AmortSbv;
       DotSbv := Methode.TableauDot[Lig];
       if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
          CessSbv := Methode.TabCession[Lig]
       else
          CessSbv := 0;

       OB.PutValue('IA_MONTANTDPI', DotSbv);  {Dotation subvention}
       OB.PutValue('IA_CESSIONDPI', CessSbv); {Cession subention}

    end;
   end;
    DatePrec := TableauDate[Lig];
    Lig := Lig + 1;
  end;
end;

procedure TPlanAmort.SetInfosCession(TypeAmort: string; DateCession: TDateTime);
begin
  TypeAmortCession := TypeAmort;
  DateAmortCession := DateCession;
end;

procedure TPlanAmort.GetCumulDotationCession(var CumulcesEco, CumulCesFis:
  double; tDateREf: TDateTime);
begin
  CumulCesEco := AmortEco.GetCumulDotationCesMet(TableauDate, tDateRef);
  CumulCesFis := AmortFisc.GetCumulDotationCesMet(TableauDate, tDateRef);

end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 04/09/2006
Modifié le ... :   /  /
Description .. : Identique procedure GetCumulDotationCession
Suite ........ : mais pour la methode PRI
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure TPlanAmort.GetCumulDotCessionPRI(var CumulCesPri:double; tDateREf: TDateTime);
begin
  CumulCesPRI := AmortPri.GetCumulDotationCesMet(TableauDate, tDateRef);
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 28/09/2006
Modifié le ... :   /  /
Description .. : Identique procedure GetCumulDotationCession
Suite ........ : mais pour la methode SBV
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure TPlanAmort.GetCumulDotCessionSBV(var CumulCesSBV:double; tDateREf: TDateTime);
begin
  CumulCesSBV := AmortSBV.GetCumulDotationCesMet(TableauDate, tDateRef);
end;


procedure TPlanAmort.ResetTableauDot(var Methode: TMethodeAmort; DateDebut:
  TDateTime);
var
  i: integer;
begin
  for i := 0 to MAX_LIGNEDOT do
  begin
    if (DateDebut > iDate1900) and (TableauDate[i] >= DateDebut) then
      Methode.TableauDot[i] := 0.00;
  end;
end;

procedure TPlanAmort.BasculeBaseCession(MethodeEco: TMethodeAmort);
var
  BaseDebut, CumulDot: double;
  i: integer;
begin
  BaseDebut := MethodeEco.Base;
  CumulDot := 0;
  i := 0;
  while MethodeEco.TableauDot[i] <> 0 do
  begin
    CumulDot := CumulDot + MethodeEco.TableauDot[i];
    MethodeEco.BaseDebutExo[i] := BaseDebut - CumulDot;
    i := i + 1;
  end;
end;

procedure TPlanAmort.CalculExceptionnelSurSortie(MethodeAmort: TMethodeAmort;
  tDateCalcul: TDateTime);
var
  Dotation, dExcEco: double;
  DateDebutEx, DateFinEx: TDateTime;
  NumDot, RestePeriode: integer;
  PremMois, PremAnnee, NbMois: Word;
begin
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  RestePeriode := GetNombrePeriodeMois(MethodeAmort.Duree - NbMoisReprise, 12);
  NumDot := 0;
  //26/03/99 saute les dotation des exercices cloturés
  while (TableauDate[NumDot] <> iDate1900) and (TableauDate[NumDot] <
    tDateCalcul) do
    inc(NumDot);

  if (RestePeriode > 0) then
  begin
    GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
    dExcEco := 0.00;
    TabTypeExc[NumDot] := ' ';
    if (DateFinEx >= DateDebutEnCours) and (DateFinEx <= DateFinEnCours) then
    begin
      GetElementExceptionnel(Dotation, dExcEco, TabTypeExc[NumDot]);
      dExcEco := dExcEco + MethodeAmort.Revision;
    end;
    MethodeAmort.TableauDot[NumDot] := MethodeAmort.TableauDot[NumDot] +
      Arrondi(dExcEco, V_PGI.OkDecV);
    TableauDate[NumDot] := DateFinEx;
    NOMBREMOIS(DateDebutEx, DateFinEx, PremMois, PremAnnee, NbMois);
    (*    if NbMois=12 then RestePeriode:=RestePeriode-1; *)
  end;
end;

procedure TPlanAmort.NettoiePlan;
var
  i: integer;
begin
  for i := 0 to MAX_LIGNEDOT do
  begin
    if AmortEco.BaseDebutExo[i] <= 0 then
    begin
      TableauDate[i] := iDate1900;
      AmortEco.TableauDot[i] := 0;
    end;
  end;
end;

// recup du dernier plan, destruction, recalcul et enregistrement :
// donc remplacement du dernier plan.

procedure TPlanAmort.UpdateBasePlan;
var
  T, TImmo: TOB;
  i, j: integer;
  numImmo, planActif: string;
  Q1: TQuery;
begin
  TImmo := TOB.Create('les immos', nil, -1);
  //  Q1:=OpenSQL('SELECT * FROM IMMO WHERE I_DATEPIECEA>="'+UsDateTime(VHImmo.EnCours.Deb)+'"',true) ;
  Q1 := OpenSQL('SELECT * FROM IMMO', true);
  TImmo.LoadDetailDB('IMMO', '', '', Q1, false);
  ferme(Q1);
  InitMove(TImmo.Detail.Count, '');
  for i := 0 to TImmo.Detail.Count - 1 do
  begin
    T := TImmo.Detail[i];
    T.PutValue('I_TYPEDEROGLIA', TypeDerogatoire (T, nil) );  // FQ 20256
    if Recupere(T.GetValue('I_IMMO'), T.GetValue('I_PLANACTIF')) then
    begin
      // recup et suppression du plan actif
      numImmo := T.GetValue('I_IMMO');
      planActif := T.GetValue('I_PLANACTIF');
      // recalcul du plan
      for j := 0 to MAX_LIGNEDOT do
        if TableauDate[j] < DateDebutEnCours then
        else
          TableauDate[j] := iDate1900;
      //InitMethodes;
      CalculTOB(T, iDate1900, fbRecalculReprise);
      NettoiePlan;
      NumSeq := StrToInt(planActif) - 1;
      ExecuteSQL('delete from immoamor where ia_immo="' + numImmo +
        '" and ia_numeroseq=' + planActif);
      SauveTOB(T);
    end;
    MoveCur(False);
  end;
  FiniMove;
  TImmo.InsertorUpdateDB(true);
  TImmo.Free;
end;

//EPZ 20/12/2000
(*

  Fonctions et procedure externes

*)

function RecalculSauvePlanTOB(OB: TOB): integer;
var
  Plan: TPlanAmort;
begin
  Plan := TPlanAmort.Create(true);
  try
    Plan.RecalculTOB(OB);
    Plan.Sauve;
    result := Plan.NumSeq;
  finally
    Plan.free;
  end;
end;

function RecalculSauvePlan(Q: TQuery): integer;
var
  Plan: TPlanAmort;
begin
  Plan := TPlanAmort.Create(true);
  try
    Plan.Recalcul(Q);
    Plan.Sauve;
    result := Plan.NumSeq;
  finally
    Plan.free;
  end;
end;

procedure RecuperePlanRepriseTheorique(Q: TQuery; var CumEco, CumFisc: double);
var
  TmpPlan: TPlanAmort;
begin
  TmpPlan := TPlanAmort.Create(true);
  try
    TmpPlan.Charge(Q);
    TmpPlan.CalculReprises;
    CumEco := TmpPlan.AmortEco.Reprise;
    CumFisc := TmpPlan.AmortFisc.Reprise;
  finally
    TmpPlan.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 29/06/2007
Modifié le ... :   /  /    
Description .. : modif de la fonction getminimulineaire
Suite ........ : si gest° Fiscale : on fait le calcul sur la durée la + longue 
Suite ........ : entre durée éco et durée fiscale
Suite ........ : sinon on fait le calcul sur la durée éco
Mots clefs ... : 
*****************************************************************}
function GetMinimumLineaire(Q: TQuery; duree :integer): double;
var
  TmpPlan: TPlanAmort;
begin
  TmpPlan := TPlanAmort.Create(true);
  try
    TmpPlan.Charge(Q);
    TmpPlan.AmortEco.Methode := 'LIN'; // On force la méthode linéaire

    // ajout mbo le 29.06.07 suite à la gestion fiscale
    TmpPlan.AmortEco.Duree := duree;

    TmpPlan.CalculReprises;
    Result := TmpPlan.AmortEco.Reprise;
  finally
    TmpPlan.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/01/2005
Modifié le ... :   /  /
Description .. : FQ 15302 - CA - 20/01/2005 - Pour lancer un recalcul des
Suite ........ : plans, il faut que les exercices comptables et immobilisations
Suite ........ : soient en phase.
Mots clefs ... :
*****************************************************************}
procedure UpdateBaseImmo(bConfirme: boolean = true; vbRecalculReprise : boolean = false);
var
  res: integer;
  msg: string;
  Plan: TPlanAmort;
begin
  { On autorise le recalcul des plans uniquement si l'exercice ouvert en immo est identique à l'exercice en cours en compta }
  if (GetParamSocSecur('SO_EXOCLOIMMO','')='') or (GetParamSocSecur('SO_EXOCLOIMMO','')=VHImmo^.Precedent.Code) then
  begin
    if bConfirme then
    begin
      msg := 'Vous allez recalculer tous les plans d''amortissement'#13#10 +
        'Souhaitez-vous poursuivre l''opération ?';
      res := PGIAsk(msg, 'Calcul des plans d''amortissement');
      if res <> mrYes then
        exit;
    end;
    Plan := TPlanAmort.Create(true);
    try
      Plan.fbRecalculReprise := vbRecalculReprise;
      if Transactions(Plan.UpdateBasePlan, 1) <> oeOK then
        MessageAlerte('Erreur lors du recalcul'#13#10'Traitement abandonné')
      else if bConfirme then
        PGIInfo('Recalcul des plans terminé.',
          'Calcul des plans d''amortissement');
    finally
      Plan.free;
    end;
  end else
    PGIBox ('Opération impossible.#10#13L''exercice comptable en cours ne correspond pas à celui des immobilisations. #10#13Veuillez effectuer la clôture comptable ou annuler la clôture des immobilisations.');
end;

function CtrlImmoOpe: boolean;
begin
  result := ExecuteSQL('select I_IMMO from IMMO where I_OPERATION="X"') = 0;
end;

function TPlanAmort.GetExcepAnter(tDebExo: TDateTime): double;
//var MontantDot : double; TypeDot : String;
var
  Q: TQuery;
  Requete: string;
begin
  result := 0.00;
  Requete := 'SELECT SUM(IL_MONTANTDOT) FROM IMMOLOG ' +
    'WHERE IL_IMMO="' + CodeImmo + '" AND IL_DATEOP<"' + USDatetime(tDebExo) +
      '"' +
    ' AND (IL_TYPEOP="ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="EEC" OR IL_TYPEOP="ELC")';
  Q := OpenSql(Requete, true);
  if not Q.Eof then
    result := Q.Fields[0].AsFloat;
  ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/08/2002
Modifié le ... : 27/04/2006
Description .. : Détermine si l'exercice de la date d'achat fait
Suite ........ :  moins de 12 mois.
Suite ........ :
Suite ........ : Modif mbo - fq 17569- date de debut n'est plus forcément la
Suite ........ : date d'achat ---> on passe en entrée la date de début
Suite ........ : d'amortissement (choisie par l'utilisateur entre date d'achat
Suite ........ : et date de mise en service)
Mots clefs ... :
*****************************************************************}

function TPlanAmort.ExoAchatInferieurA12(DateDeb:TdateTime): boolean;
var
  Exo: string;
  ExoDt: TImExoDate;
  PremMois, PremAnnee, nMois: Word;
begin
  Result := False;
  // mbo fq 17569 Exo := ImQuelExoDt(DateAchat);
  Exo := ImQuelExoDt(DateDeb);
  if Exo <> '' then
  begin
    if ImQuelDateDeExo(Exo, ExoDt) then
    begin
      ImNOMBREPEREXO(ExoDt, PremMois, PremAnnee, nMois);
      Result := nMois < 12;
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 23/03/2006
Modifié le ... :   /  /
Description .. : Détermine si l'exercice dont la date debut est passée en
Suite ........ : paramètre fait moins de 12 mois (utilisé pour le calcul en
Suite ........ : mode dégressif)
Mots clefs ... :
*****************************************************************}
function TPlanAmort.ExoInferieurA12(DateDebEx:TdateTime): boolean;
var
  Exo: string;
  ExoDt: TImExoDate;
  PremMois, PremAnnee, nMois: Word;
begin
  Result := False;
  Exo := ImQuelExoDt(DateDebEx);
  if Exo <> '' then
  begin
    if ImQuelDateDeExo(Exo, ExoDt) then
    begin
      ImNOMBREPEREXO(ExoDt, PremMois, PremAnnee, nMois);
      Result := nMois < 12;
    end;
  end;
end;


procedure TPlanAmort.ChargePlanUO(T: TOB);
begin
  fPlanUO.Dupliquer(T, True, True);
end;

function TPlanAmort.DateDebutEnCours: TDateTime;
begin
{$IFDEF EAGLSERVER}
  Result := fEncours.Deb;
{$ELSE}
  Result := VHImmo^.Encours.Deb;
{$ENDIF}
end;

function TPlanAmort.DateFinEnCours: TDateTime;
begin
{$IFDEF EAGLSERVER}
  Result := fEncours.Fin;
{$ELSE}
  Result := VHImmo^.Encours.Fin;
{$ENDIF}
end;

{$IFDEF EAGLSERVER}

procedure TPlanAmort.SetDatesServer(Encours: TExoDate);
begin
  fEncours := Encours;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Calcul de plan d'amortissement pour la sortie
modif          : mbo 02/10/2007 ajout du paramètre calculPourSortie uniquement à true lors d'une cession
                 (false lors d'un remplacement)
                 mbo 29/10/2007 - fq 21754 - ajout paramètre amortJsortie pour prise en cpte ou non jour de cession
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.TraiteSortie(DateOperation: TDateTime; ModeCalcul : string;
                                  MntExcSortie : double; stTypeExc : string;
                                  AmortJsortie : string = '');
var
  i, j: integer;
  Prorata : double;
  CessionZero,CessionMoinsUn : boolean; //new
begin
  i := 0;

  { Bascule des informations sur les champs de sortie }
  while ((TableauDate[i] <> iDate1900) and (TableauDate[i] < VHImmo^.Encours.Fin)) do

  begin
    AmortEco.BaseDebExoCes[i] := AmortEco.Base;
    AmortEco.TabCession[i] := AmortEco.TableauDot[i];
    AmortEco.TableauDot[i] := 0;
    AmortFisc.TabCession[i] := AmortFisc.TableauDot[i];
    AmortFisc.TableauDot[i] := 0;

    AmortPri.TabCession[i] := AmortPri.TableauDot[i];
    AmortPri.TableauDot[i] := 0;
    AmortSbv.TabCession[i] := AmortSbv.TableauDot[i];
    AmortSbv.TableauDot[i] := 0;

    i := i + 1;
  end;
  { Mise à jour du plan de l'année concernant la sortie }
  // MntExcSortie = exceptionnel saisi au moment de la cession
  //  FQ 16711 - MontantExc   = dotation exceptionnelle qui ne doit pas être proratisée

  if (stTypeExc = 'RDO') then MntExcSortie := (-1)*MntExcSortie;

  if (TableauDate[i] = VHImmo^.Encours.Fin) then
  begin
    AmortEco.BaseDebExoCes[i] := AmortEco.Base;
    if ModeCalcul = 'SAN' then
    begin
      AmortEco.TabCession[i] := MntExcSortie;
      AmortFisc.TabCession[i] := 0;

      // pour calcul dérog et reintegration chantier fiscal
      AmortEco.TableauDot[i] := AmortEco.TabCession[i];
      AmortFisc.TableauDot[i] := AmortFisc.TabCession[i];

    end else
    begin
      //ajout mbo - fq 16711 - 20/09/2005 -
      if (TypeExc = 'RDO') then MontantExc := (-1)*MontantExc;

      // fq 21754 - mbo - 29/10/2007
      if (DateOperation = DateDebutEnCours) and (AmortJsortie = 'NON') then
      begin
         CessionMoinsUn := false; //ds ce cas ce paramètre ne sert pas
         CessionZero := true     // cession = 1er jour de l'exercice et pas d'amortissement jour de cession
      end else
      begin
         CessionZero := false;
         if (AmortJsortie = 'NON') then
             CessionMoinsUn := true    // pas d'amortissement du jour de cession
         else
             CessionMoinsUn := false;   // amortissement du jour de cession
      end;
      { Economique }
      if (AmortEco.Methode = 'LIN') then
        // fq 16711 - AmortEco.TabCession[i] := GetDotationLinAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortEco), DateOperation, AmortEco.DateFinAmort, AmortEco.TableauDot[i], False)+MntExcSortie
      begin
         if (CessionZero) then
             AmortEco.TabCession[i] := (MntExcSortie+MontantExc)
         else if (CessionMoinsUn) then
             AmortEco.TabCession[i] := GetDotationLinSortie(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortEco), DateOperation, AmortEco.DateFinAmort, (AmortEco.TableauDot[i]-MontantExc), False)+MntExcSortie+MontantExc
         else
             AmortEco.TabCession[i] := GetDotationLinAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortEco), DateOperation, AmortEco.DateFinAmort, (AmortEco.TableauDot[i]-MontantExc), False)+MntExcSortie+MontantExc;

      end else
      begin
         if AmortEco.Methode = 'VAR' then
            AmortEco.TabCession[i] := GetDotationLinAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortEco), DateOperation, AmortEco.DateFinAmort, (AmortEco.TableauDot[i]-MontantExc), False)+MntExcSortie+MontantExc

         else if (AmortEco.Methode = 'DEG') then
            // fq 16711 - AmortEco.TabCession[i] := GetDotationDegAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortEco), DateOperation, AmortEco.TableauDot[i], Prorata, Suramort)+MntExc
            AmortEco.TabCession[i] := GetDotationDegAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortEco), DateOperation, (AmortEco.TableauDot[i]-MontantExc), Prorata, Suramort)+MntExcSortie+MontantExc
         else AmortEco.TabCession[i] := 0;
      end;

      { Fiscal }
      // fq 17102 - mbo 29/11/05 - ajout de l'exceptionnel
      // fq 17154 - mbo 06.12.2005 - excep = pas impact sur le fiscal pour immo plan VAR
      if (AmortEco.Methode = 'VAR') then
      begin
         // FQ 21546 - le fiscal peut être linéaire ou dégressif
         if AmortFisc.Methode = 'LIN' then
            AmortFisc.TabCession[i] := GetDotationLinAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortFisc), DateOperation, AmortEco.DateFinAmort, AmortFisc.TableauDot[i], False)
         else
            AmortFisc.TabCession[i] := GetDotationDegAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortFisc), DateOperation, AmortFisc.TableauDot[i]-MontantExc, Prorata, Suramort);
      end else
      begin
        if (AmortFisc.Methode = 'LIN') then
          // correction bug mbo - 13.04.2007 AmortFisc.TabCession[i] := GetDotationLinAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortFisc), DateOperation, AmortEco.DateFinAmort, AmortFisc.TableauDot[i]-MontantExc, False)+MntExcSortie+MontantExc
        begin
             if CessionZero then
                AmortFisc.TabCession[i] := MntExcSortie+MontantExc
             else if CessionMoinsUn then
                AmortFisc.TabCession[i] := GetDotationLinSortie(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortFisc), DateOperation, AmortFisc.DateFinAmort, AmortFisc.TableauDot[i]-MontantExc, False)+MntExcSortie+MontantExc
             else
                AmortFisc.TabCession[i] := GetDotationLinAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortFisc), DateOperation, AmortFisc.DateFinAmort, AmortFisc.TableauDot[i]-MontantExc, False)+MntExcSortie+MontantExc;
        end
        else if (AmortFisc.Methode = 'DEG') then
          AmortFisc.TabCession[i] := GetDotationDegAvant(VHImmo^.EnCours.Deb, VHImmo^.EnCours.Fin, GetDateDebutAmort(AmortFisc), DateOperation, AmortFisc.TableauDot[i]-MontantExc, Prorata, Suramort)+MntExcSortie+MontantExc
        else AmortFisc.TabCession[i] := 0;
      end;

      AmortPri.TabCession[i] := GetRestePri(AmortPri, VHImmo^.Encours.Deb);
      AmortSbv.TabCession[i] := GetResteSbv(AmortSbv, VHImmo^.Encours.Deb);

    end;


    AmortDerog.TableauDot[i] := 0;
    AmortReint.TableauDot[i] := 0;


    AmortEco.TableauDot[i] := 0;
    AmortFisc.TableauDot[i] := 0;

    AmortPri.TableauDot[i] := 0;
    AmortSbv.TableauDot[i] := 0;

    { On se positionne sur l'annuité suivante }
    Inc(i);
  end;
  { Mise à 0 des informations pour les dates postérieures à la sortie }
  for j := i to MAX_LIGNEDOT do
  begin
    TableauDate[j] := iDate1900;
    AmortEco.BaseDebExoCes[j] := 0;
    AmortEco.TabCession[j] := 0;
    AmortEco.TableauDot[j] := 0;
    AmortFisc.TabCession[j] := 0;
    AmortFisc.TableauDot[j] := 0;

    AmortPri.TabCession[j] := 0;
    AmortPri.TableauDot[j] := 0;
    AmortSbv.TabCession[j] := 0;
    AmortSbv.TableauDot[j] := 0;

    // ajout pour chantier fiscal
    AmortDerog.TableauDot[j] := 0;
    AmortDerog.TabCession[j] := 0;
    AmortReint.TableauDot[j] := 0;
    AmortReint.TabCession[j] := 0;

  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 27/04/2006
Modifié le ... : 04/09/2006
Description .. : Modif pour fq 17569
Suite ........ : Ajout du plan en entrée : 'FISC' ou 'ECO'
Suite ........ : si plan éco = date = datedebeco
Suite ........ : sinon  date = datedebfis
Suite ........ : modif mbo le 04.09.06 : ajout du plan PRI
Suite ........ : modif mbo le 29.09.06 : ajout du plan SBV
Mots clefs ... :
*****************************************************************}
function TPlanAmort.CalculDureeReprise(PlanUtil: string): integer;
var
    PremMois, PremAnnee, NbMois: Word;
begin
  if PlanUtil = 'ECO' then
     NOMBREMOIS(DateDebEco, DateDebutEnCours, PremMois, PremAnnee, NbMois)
  else if PlanUtil = 'FISC' then
     NOMBREMOIS(DateDebFis, DateDebutEnCours, PremMois, PremAnnee, NbMois)
  else if PlanUtil = 'PRI' then
     NOMBREMOIS(PRIDateDeb, DateDebutEnCours, PremMois, PremAnnee, NbMois)
  else   //sbv
     NOMBREMOIS(SBVDateDeb, DateDebutEnCours, PremMois, PremAnnee, NbMois);

  Result := MaxIntValue([0, NbMois - 1]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 10/11/2005
Modifié le ... :   /  /
Description .. : Quand la dotation en dégressif passe en linéaire, le calcul
Suite ........ : de la reprise maximale de dépréciation doit suivre le même
Suite ........ : principe.
Suite ........ : Cette fonction renvoie le taux à appliquer pour calculer la
Suite ........ : reprise
Mots clefs ... :
*****************************************************************}
function TPlanAmort.DegressifTauxReprise(MethodeAmort: TMethodeAmort; tDateCalcul:
  TDateTime; bCloture: boolean): double;
var
  EvalLineaire, Prorata, Dotation, Cumul, Base, Taux: double;
  SommeDot: double;
  DateRef, DateDebutEx, DateFinEx, wDate: TDateTime;
  NumDot, RestePeriode: integer;
  //PremMois,PremAnnee,NbMois : Word;
begin
  { mbo - 25.08.06
    suite correction pour conseil compil, datedebamort n'est plus utilisé

  //fq 17569 ajout mbo -
  If MethodeAmort.PlanUtil = 'FISC' then
     DateDebAmort := DateDebFis
  else
     DateDebAmort := DateDebEco;
  }
  Base := MethodeAmort.Base;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  Taux := MethodeAmort.Taux;
  NumDot := 0;
  SommeDot := 0.00;
  Cumul := MethodeAmort.Reprise;
  if ImmoCloture then
    MethodeAmort.TraiteDotationCloture(TableauDate, NumDot, SommeDot, bCloture);

  Cumul := Cumul + SommeDot;
  Vnc := Base - Cumul;
  //Calcul du nombre de periodes restantes
  wDate := tDateCalcul;
  RestePeriode := 0;
  while wDate <= MethodeAmort.DateFinAmort do
  begin
    inc(RestePeriode);
    GetDatesExercice(wDate, DateDebutEx, DateFinEx);
    wDate := DateFinEx + 1;
  end;
  // Si exercice d'achat inférieur à 12 mois, on ajoute une période.
  // if ExoAchatInferieurA12 then Inc (RestePeriode); CA - 02/04/2004 - Inutile car déjà pris en compte dans la date de fin d'amortissement
  //Pour gérer l'anteriorité non-modifiable
  while (TableauDate[NumDot] <> iDate1900) and (TableauDate[NumDot] <
    tDateCalcul) do
  begin
    Cumul := Cumul + MethodeAmort.TableauDot[NumDot];
    inc(NumDot);
    Vnc := Base - Cumul;
  end;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  if Arrondi(Vnc, V_PGI.OkDecV) > 0.00 then
  begin
    GetDatesExercice(tDateCalcul, DateDebutEx, DateFinEx);
    DateRef := DateDebutEx;

    // fq 17569 if (DateAchat > DateDebutEx) and (DateAchat <= DateFinEx) then
    // DateRef := DateAchat;

    { mbo modif conseil de compil 25.08.06 - test inutile car dotation non exploitée
    if (DateDebAmort > DateDebutEx) and (DateDebAmort <= DateFinEx) then
       Dotation := Vnc;
    }

    if MethodeAmort.DateFinAmort > DateFinEx then
    begin
      Dotation := GetDotationDegApres(DateDebutEx, DateFinEx, DateRef, (*Base*)
        Vnc * Taux, Prorata);
      if (RestePeriode > 0) then
      begin
        EvalLineaire := ((Vnc / RestePeriode) * Prorata);
        if (EvalLineaire > Dotation) then
          taux := (Prorata/RestePeriode);
      end else
         taux := 1;
    end else
      taux := 1;
  end;
  result := taux;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 10/02/2006
Modifié le ... :   /  /
Description .. : calcul des antérieurs théoriques sans tenir compte d'une
Suite ........ : saisie d'antérieur dépréciation
Mots clefs ... :
*****************************************************************}
procedure TPlanAmort.CalculRepriseTheorique(MethodeAmort: TMethodeAmort; tDateCalcul:TDateTime);
begin
  CalculDateFinAmortReprise(MethodeAmort);
  MethodeAmort.Reprise := 0.00;
  if GetDateDebutAmort(MethodeAmort) > tDateCalcul then
    exit;
  if MethodeAmort.Methode = 'LIN' then
    CalculRepriseLineaire(MethodeAmort, GetDateDebutAmort(MethodeAmort),
      tDateCalcul)
  else if MethodeAmort.Methode = 'DEG' then
    CalculRepriseDegressif(MethodeAmort, tDateCalcul)
  else if MethodeAmort.Methode = 'VAR' then
    CalculRepriseVariable(MethodeAmort, tDateCalcul);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 10/02/2006
Modifié le ... :   /  /
Description .. : fonction permettant de recalculer le montant de la reprise
Suite ........ : maximum dépréciation à partir de l'antérieur dépréciation
Suite ........ : saisi lors de la création de l'immo
Suite......... : la formule =

                 etape 1  = antérieur dépréciation -
                               (antérieur théorique - antérieur réél)
                 etape 2  = on applique à étape 1 un coefficient calculé suivant la méthode d'amort.

Mots clefs ... :
*****************************************************************}

function TPlanAmort.CalculRepriseDepreciation(MethodeAmort: TMethodeAmort;
  tDateDeb, tDateFin: TDateTime; Clot: boolean): double;
var QLog: TQuery;
    type_op : string;
    Cumul : Double;
    nb1 : Integer;
    nb2 : Integer;
    Prorata:Double;
    Montant:Double;
    TauxCalcule:Double;
    AnterieurTheorique:Double;
    AnterieurReel:Double;
    CumulAnt:Double;
    //dCumulEco:Double;
    dCumulFisc:Double;
    Methode2 : TMethodeAmort;
    tabdat2 :array[0..MAX_LIGNEDOT] of TDateTime;
    i: integer;
begin
Cumul :=0.00;

AnterieurReel:=0.00;
//dCumulEco:=0.00;
CumulAnt:=0.00;
Montant:=0.00;

if MethodeAmort.RepriseDep <> 0 then
begin
  // on sauvegarde le tableau des dates du plan amort en cours
  // car ce tableau date va être modifié par le calcul des antérieurs théoriques
  for i := low(TableauDate) to high(TableauDate) do
      TabDat2[i] := TableauDate[i];

  Methode2:=TMethodeAmort.Create;
  methode2.Copie(MethodeAmort,'ECO');;
  Methode2.Reprise := 0;
  Methode2.RepriseDep := 0;
  // on initialise le tableau des dot de la nouvelle methode sinon le calcul en dégressif est faux
  for i:=low(Methode2.TableauDot) to high(Methode2.TableauDot) do Methode2.TableauDot[i]:=0.00;

  CalculRepriseTheorique(Methode2,tDateDeb);

  AnterieurTheorique :=Methode2.Reprise;
  Methode2.free;

  // on restitue le tableau des dates du plan en cours
  for i := low(TabDat2) to high(TabDat2) do
      TableauDate[i] := TabDat2[i];

  GetCumulsDotExercice(tDateDeb,AnterieurReel,dCumulFisc,false,true,false);

  //getCumulsDotExercice renvoie antérieur eco + Anterieur dépréciation
  AnterieurReel := AnterieurReel - MethodeAmort.RepriseDep;
  CumulAnt:= MethodeAmort.RepriseDep - (AnterieurTheorique - AnterieurReel);
end;

if clot = false then
   Cumul := CumulAnt
else
begin
  type_op:='CLO';

  QLog := OpenSQL('SELECT IL_REVISIONECO FROM IMMOLOG WHERE (IL_IMMO="'
      + CodeImmo + '" AND IL_TYPEOP ="' + type_op +'" AND IL_DATEOP<"' +
      USDateTime(VHImmo^.Suivant.Deb) +'")' , True);

  while not QLog.Eof do
  begin
      Cumul :=  QLog.FindField('IL_REVISIONECO').AsFloat;
      QLog.Next;
  end;
  Ferme(QLog);

  if (Cumul = 0) and (CumulAnt <> 0) then
     Cumul := CumulAnt
  else
     if (cumul = 0) and (cumulAnt = 0) then
        cumul := RevisionEco;

  // Cumul := Cumul + RevisionEco; // on ajoute la dépréciation saisie sur l'exercice que l'on cloture

end;


  if MethodeAmort.Methode = 'LIN' then
  begin
    //nb1 = nb jours de l'exercice - nb2 = nb jours restant à amortir
    nb1 := NOMBREJOUR360 (tDateDeb,tDateFin);
    nb2 := NOMBREJOUR360 (tDateDeb, MethodeAmort.DateFinAmort);

    if nb1 > nb2 then   // cas de la dernière année
      montant := 0.00
    else
      montant := cumul * (1 - (nb1/nb2));

  end else
    if MethodeAmort.Methode = 'VAR' then
    begin
       prorata:= ProrataUOreste(tDateDeb);
       montant:= cumul * (1-prorata);
    end else

    if MethodeAmort.methode = 'DEG' then
    begin
       TauxCalcule := DegressifTauxReprise(MethodeAmort,tDateDeb, true);
       montant := cumul * (1-TauxCalcule);
  end;


  // on ajoute la depreciation saisie sur l'exercice
  //Montant := Montant + RevisionEco;
  Result := arrondi(montant,V_PGI.OkDecV);


end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Détermination si gestion fiscale ou non
Suite ........ : Dans le cadre des évolutions fiscales : le but de cette
Suite ........ : fonction = renvoyer true si la première dotation éco calculée
Suite ........ : est supérieure à la dotation fiscale calculée (à périmètre
Suite ........ : égal)
Suite ........ : ATTENTION : dans le cadre d'un éco variable : ne pas appeler cette
                 fonction : c'est le client qui décide

Mots clefs ... :
*****************************************************************}

Function TplanAMort.Determine_gestion_fiscale(Q: TQuery; RepriseEco, RepriseFisc:double ):boolean;
Var
  tabdat2 :array[0..MAX_LIGNEDOT] of TDateTime;
  i: integer;
  savDateDebEco : TdateTime;
  savDateDebFis : TdateTime;
  AnterieurTheoriqueE :double;
  AnterieurTheoriqueF :double;
  IndiceExoFisc, IndiceExoEco : integer;
  Plan : tPlanAmort;

begin
IndiceExoFisc := 0;
IndiceExoEco := 0;

Plan := TPlanAmort.Create(true);
try
  Plan.Charge(Q);

  if plan.Fiscal = false then
     result := false  // si pas de plan fiscal

  else if RepriseEco > RepriseFisc then
     result := true //gestion fiscale systématique

  //calculer dotation éco et fiscale sur une durée théorique (en l'occurence
  // date début d'exercice en cours à date fin d'exercice en cours).

  // on sauvegarde le tableau des dates du plan amort en cours
  // car ce tableau date va être modifié par le calcul des antérieurs théoriques
  else
  begin
    for i := low(TableauDate) to high(TableauDate) do
      TabDat2[i] := TableauDate[i];

      Calcul(Q, iDate1900);

      for i := low(TableauDate) to high(TableauDate) do
      begin
         if TableauDate[i] >= DateDebfis then
         begin
            indiceexofisc := i;
            Break;
         end;
      end;

      for i := low(TableauDate) to high(TableauDate) do
      begin
        if TableauDate[i] >= DateDebEco then
        begin
            indiceexoEco := i;
            Break;
        end;
      end;

      if indiceExoEco > indiceExoFisc then
         result := false
      else if indiceExoEco < indiceExoFisc then
         result := true
      else
      begin
         // on initialise le tableau des dates
           for i := 0 to MAX_LIGNEDOT do
              TableauDate[i] := iDate1900;

         savDateDebEco := DateDebEco;
         DateDebEco := VHImmo^.Encours.deb;
         AmortEco.Reprise := 0.00;
         AmortEco.RepriseDep := 0.00;
         // on initialise le tableau des dot
         for i:=low(AmortEco.TableauDot) to high(AmortEco.TableauDot) do AmortEco.TableauDot[i]:=0.00;

         CalculDotation(AmortEco, VHImmo^.Encours.fin, false);
         AnterieurTheoriqueE :=AmortEco.TableauDot[0];

         DateDebEco := SavDateDebEco;

         // calcul antérieur théorique fiscal
         savDateDebFis := DateDebFis;
         DateDebFis := VHImmo^.Encours.deb;
         AmortFisc.Reprise := 0.00;
         AmortFisc.RepriseDep := 0.00;

         for i:=low(AmortFisc.TableauDot) to high(AmortFisc.TableauDot) do AmortFisc.TableauDot[i]:=0.00;

         CalculDotation(AmortFisc, VHImmo^.Encours.Fin, false);
         AnterieurTheoriqueF :=AmortFisc.TableauDot[0];

         DateDebFis := savDateDebFis;

         if AnterieurTheoriqueE >= AnterieurTheoriqueF then
            result := true
         else
            result := false;

           // on restitue le tableau des dates du plan en cours
         for i := low(TabDat2) to high(TabDat2) do
              TableauDate[i] := TabDat2[i];

      end;
  end;
finally
  Plan.free;
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 20/02/2007
Modifié le ... :   /  /
Description .. : Procédure qui calcule le dérogatoire ou la réintégration fiscale
Suite ........ : et alimente les 2 méthodes DRG ou REIN
Mots clefs ... :
*****************************************************************}
Procedure TplanAMort.Calcul_Derog_Reint;
Var
  //DatePrec : TDateTime;
  DateFinAmort: TDateTime;
  Lig: integer;
  //CessEco, CessFisc : double
  CumulEco, CumulFisc, DotDerog, DotEco, DotFisc: double;

  //Dérogatoire
  CE, CF, AE, AF, EC, EA : double;


  DateDebut, DateFin : TDateTime;

  DotReint : double;
  cumulreint : double;
  cumulderog : double;

begin
//DatePrec := iDate1900;

DateDebut := iDate1900;
DateFin   := iDate1900;


try
  Modifie := false;
  CumulEco := arrondi((AmortEco.Reprise + AmortEco.RepriseCedee),V_PGI.OkDecV);
  CumulFisc := arrondi((AmortFisc.Reprise + AmortFisc.RepriseCedee),V_PGI.OkDecV);

  DateFinAmort := GetDateFinAmort;    // fonction modifiée pour prime et sbv
  // ajout mbo pour éviter les lignes à 0 en fin d'immoamor
  // recherche fin exercice de l'exercice de fin d'amortissement
  GetDatesExercice(DateFinAmort, DateDebut, DateFin);

  Lig := 0;

  CE := arrondi((AmortEco.Reprise + AmortEco.RepriseCedee),V_PGI.OkDecV);
  CF := arrondi((AmortFisc.Reprise + AmortEco.RepriseCedee),V_PGI.OkDecV);

  if AmortEco.Methode = 'VAR' then
     CE := arrondi((CE + AmortEco.RepriseDep),V_PGI.OkDecV);

  CumulDerog := arrondi((AmortDerog.Reprise + AmortDerog.RepriseCedee),V_PGI.OkDecV);
  CumulReint := arrondi((AmortReint.Reprise + AmortReint.RepriseCedee),V_PGI.OkDecV);

  AE := CE;
  AF := CF;

  while TableauDate[Lig] <> iDate1900 do
  begin

    if TableauDate[Lig] <> iDate1900 then
    begin
      //Dotation économique
      DotEco := arrondi((AmortEco.TableauDot[Lig] + AmortEco.TabCession[Lig]),V_PGI.OkDecV);
      //Dotation partie(s) cédée(s)
      //CessEco := 0.00;
      {
      if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
          CessEco := AmortEco.TabCession[Lig];
      }
      DotFisc := 0.00;
      DotDerog := 0.00;
      DotReint := 0.00;
      //CessFisc := 0.00;

      if AmortFisc.Methode <> '' then
      begin
        DotFisc := arrondi((AmortFisc.TableauDot[Lig] + AmortFisc.TabCession[Lig]),V_PGI.OkDecV);
        {
        if (TableauDate[Lig] <= DateFinAmort) or (TableauDate[Lig] > DatePrec) then
            CessFisc := AmortFisc.TabCession[Lig];
        }
        // Dérogatoire
        CE := arrondi((CE + DotEco),V_PGI.OkDecV);
        CF := arrondi((CF + DotFisc),V_PGI.OkDecV);

        EC := arrondi((CF-CE),V_PGI.OkDecV); //écart sur (cumul antérieur + dotation de l'exercice)
        EA := arrondi((AF-AE),V_PGI.OkDecV); //écart sur cumuls antérieurs

        DotDerog := 0;
        DotReint := 0;

        if AmortEco.Methode = 'VAR' then
        begin

          if not (GestionFiscale) then
          begin
            // on fait les calculs comme avant
            if (EC >= 0) then     // écart sur antérieurs
            begin
              if EA < 0 then EA := 0;
              if EC > EA then DotDerog := arrondi((EC-EA),V_PGI.OkDecV)
              else if EA > EC then DotDerog := arrondi((EC-EA),V_PGI.OkDecV);
            end else
              if (EC < 0) then if (EA>0) then DotDerog := (-1)*EA;

          end else
          begin

             // ajout chantier fiscal : on solde le dérogatoire avant de générer la réintégration
             // mais pas l'inverse : c'est à dire qu'on ne se préoccupe pas du cumul réintégration avant de
             // générer du dérogatoire

             if (DotFisc > DotEco) then
             begin
               DotDerog := arrondi((DotFisc - DotEco),V_PGI.OkDecV);
               CumulDerog := arrondi((CumulDerog + DotDerog),V_PGI.OkDecV);
               DotReint := 0;
             end else
             begin
               if CumulDerog = 0 then
               begin
                  DotReint := (DotFisc - DotEco) * -1;
                  DotReint := arrondi(DotReint,V_PGI.OkDecV);
                  CumulReint := arrondi((CumulReint + DotReint),V_PGI.OkDecV);
               end else
               begin
                 if CumulDerog > arrondi(((DotFisc - DotEco)* -1),V_PGI.OkDecV) then
                 begin
                   DotDerog := arrondi((DotFisc - DotEco),V_PGI.OkDecV);  // on est en reprise de dérogatoire
                   CumulDerog := arrondi((CumulDerog + DotDerog),V_PGI.OkDecV);
                 end else
                 begin
                   // on va solder le dérog et générer de la réintégration
                   DotDerog := CumulDerog * -1;
                   DotReint := (CumulDerog + DotFisc - DotEco)*-1;
                   DotDerog := arrondi(DotDerog,V_PGI.OkDecV);
                   DotReint := arrondi(DotReint,V_PGI.OkDecV);
                   CumulDerog := 0;
                   CumulReint := arrondi((CumulReint + DotReint),V_PGI.OkDecV);
                 end;
               end;
             end;
          end;
        end else
        begin
          // le plan d'amortissement éco n'est pas variable
          if not (GestionFiscale) then
          begin
            // on fait les calculs comme avant
            if (EC >= 0) then     // écart sur antérieurs
            begin
              if EA < 0 then EA := 0;
              if EC > EA then DotDerog := arrondi((EC-EA),V_PGI.OkDecV)
              else if EA > EC then DotDerog := arrondi((EC-EA),V_PGI.OkDecV);
            end else
              if (EC < 0) then if (EA>0) then DotDerog := (-1)*EA;

          end else  // traitement de la gestion fiscale
          begin
             // ajout chantier fiscal : on génére une réintégration ou une déduction
             // qd c'est positif = réintégration - qd c'est négatif = déduction
             DotReint := arrondi((DotEco - DotFisc),V_PGI.OkDecV);
             CumulReint := arrondi((CumulReint + DotReint),V_PGI.OkDecV);
             DotDerog := 0;
          end;
        end;

        AE := CE;
        AF := CF;
      end;

      //Dotation Dérogatoire
      AmortDerog.TableauDot[Lig] := DotDerog;
      // réintégration extra comptable
      AmortReint.TableauDot[Lig] := DotReint;

      CumulEco := arrondi((CumulEco + DotEco),V_PGI.OkDecV);
      if ((AmortFisc.Methode) <> '') then
        CumulFisc := arrondi((CumulFisc + DotFisc),V_PGI.OkDecV);

    end;
    //DatePrec := TableauDate[Lig];
    Lig := Lig + 1;
  end;
finally
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Calcul le cumul dérogatoire et cumul réintégration à une
Suite ........ : date donnée :
Suite ........ : utilisée lors de la cession pour cumul à fin d'exercice N-1
Suite ........ : ds ce cas appel avec date deb d'exercice N
Mots clefs ... :
*****************************************************************}
Procedure TPlanAmort.CumulDerogReint(tDateCalcul:TdateTime;var CumDerog, CumReint: double);
var
  i : integer;

begin

  i := 0;

  CumDerog := arrondi((AmortDerog.Reprise + AmortDerog.RepriseCedee),V_PGI.OkDecV); // les antérieurs saisis
  CumReint := arrondi((AmortReint.Reprise + AmortReint.RepriseCedee),V_PGI.OkDecV);

  while ((TableauDate[i]<>iDate1900) and (TableauDate[i]<tDateCalcul)) do
  begin
    CumDerog := CumDerog + AmortDerog.TableauDot[i] +AmortDerog.TabCession[i];
    CumReint := CumReint + AmortReint.TableauDot[i] +AmortReint.TabCession[i];
    Inc (i);
  end;

  CumDerog := arrondi(CumDerog,V_PGI.OkDecV);
  CumReint := arrondi(CumReint,V_PGI.OkDecV);

end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 22/03/2007
Modifié le ... :   /  /
Description .. : Cette fonction est appelée quand on recalcule les reprises
Suite ........ : (appelée par la fonction CalculReprises)
Mots clefs ... :
*****************************************************************}
Procedure TplanAmort.CalculReprise_Derog_Reint;
var
  CumAntEco, CumAntFisc : double;


begin
  if AmortEco.Methode = 'VAR' then
     CumAntEco := AmortEco.Reprise + AmortEco.RepriseDep
  else
     CumAntEco := AmortEco.Reprise;

  CumAntFisc:= AmortFisc.Reprise;

  if not(GestionFiscale) then
  begin
     if CumAntFisc< CumAntEco then
     begin
        AmortDerog.Reprise := 0;
        AmortReint.Reprise := 0;
     end else
     begin
        AmortReint.Reprise := 0;
        AmortDerog.Reprise := (CumAntFisc - CumAntEco);
     end;

  end else
  begin
     begin
        if AmortEco.Methode = 'VAR' then
        begin
           if (CumAntFisc < CumAntEco) then
           begin
              AmortDerog.reprise := 0;
              AmortReint.reprise := (CumAntEco - CumAntFisc);
           end else
           begin
              AmortDerog.Reprise := (CumAntFisc - CumAntEco);
              AmortReint.Reprise := 0;
           end;
        end else
        begin
           // on ne gère plus de dérogatoire : tout est en réintégration/Déduction
           if (CumAntFisc< CumAntEco) then
           begin
              AmortDerog.Reprise := 0;
              AmortReint.reprise := (CumAntEco - CumAntFisc);
           end else
           begin
              AmortDerog.Reprise := 0;
              AmortReint.reprise := (CumAntEco - CumAntFisc);
           end;
        end;
     end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Calcul de la valeur résiduelle comptable ou fiscale suivant la
Suite ........ : méthode transmise
                 la valeur résiduelle se calcule par rapport à la base amortissable
                 (la valeur nette se calcule par rapport à la valeur d'achat)
Mots clefs ... : Fonction utilisée pour le calcul du cumul de la réintégration sur l'immo
                 remplacée
*****************************************************************}
function TPlanAmort.GetValResiduelle(CodeImmo: string; MethAmort:TMethodeAmort; DateRef:
  TDateTime; bCession: boolean): double;
var
  CumulDot, base, anterieurs: double;
  Indice: integer;
begin
  CumulDot := 0.00;
  Indice := 0;

  if bCession then
  begin
     while (TableauDate[Indice] <> iDate1900) do
     begin
        CumulDot := CumulDot + MethAmort.TabCession[Indice];
        Indice := Indice + 1;
     end;

  end else
  begin
     while (TableauDate[Indice] <> iDate1900) and (TableauDate[Indice] <= DateRef) do
     begin
      CumulDot := CumulDot + MethAmort.TableauDot[Indice];
      Indice := Indice + 1;
     end;
  end;

  if MethAmort.PlanUtil = 'ECO' then
     base := Recup_Baseeco(Codeimmo, bCession)
  else
     base := Recup_BaseFisc(CodeImmo, bCession);

  if bCession then
     Anterieurs := MethAmort.RepriseCedee + MethAmort.RepriseDep
  else
     Anterieurs := MethAmort.Reprise + MethAmort.RepriseDep;


  Result := Base - (CumulDot+Anterieurs);

end;


end.




