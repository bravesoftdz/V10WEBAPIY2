{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/06/2003
Modifié le ... : 05/08/2004
Description .. : CA - 02/06/2003 - FQ 12370 : calcul VNC avec
Suite ........ : exceptionnel
Suite ........ : CA - 02/06/2003 - FQ 12369 : non proratisation de
Suite ........ : l'exceptionnel
Suite ........ : CA - 20/01/2004 - FQ 13217 : calcul réintégration si sortie
Suite ........ : CA - 05/08/2004 - FQ 14093 : mise à jour de la base
Suite ........ : économique dans la cas d'une immobilisation non
Suite ........ : amortissable
Suite ........ : MBO - 30/09/2005 - FQ 16803 : état prévisionnel : dot fiscale fausse si dérog=0
Suite ........ : MBO - 30/09/2005 - FQ 15280 : gestion amort exceptionnel sur immos avec dérog.
Suite......... : MBO - 24/10/2005 - Modif appel GetExcepExercice pour dépréciation actif
Suite......... : MBO - 16/11/2005 - ajout de dTotExc = total exeptionnel+dépréciation actif
                                    dans la classe tplaninfo
Suite......... : MBO - 18/11/2005 - ajout fonction GetDotDeprec = recupere dépréciation d'actif
                                    en fonction d'une date donnée
Suite......... : MBO - 30/11/2005 - FQ 17013 ajout dTotExcFisc et dTotExcEco pour cumul exceptionnel
Suite......... : MBO - 10/01/2006 - FQ 17289 VNC négative en édition tableau éco + liste simplifiée
                                    si immo cédée et impression avant date de cession
Suite......... : MBO - 10/01/2006 - FQ 17292 Valeur HT fausse en édition tableau éco si immo cédée
                                    (non prise en compte de tva récupérable et tva récupérée)
Suite......... : MBO - 02/06      - FQ 17476 - impact de la dépréciation sur le plan fiscal
Suite......... : BTY - 02/06      - Ouverture des pgms d'édition à Web Access
Suite......... : BTY - 02/06 FQ 17547 - Montants faux en Web Access cause mauvaise gestion des dates sur le serveur (indice 0)
Suite......... : BTY - 03/06 Pb en éditions Web Access ((éditions ne sortent pas
Suite......... : ou certains montants ne sortent pas ou sont faux) à cause des dates lues
Suite......... : dans VHIMMO au lieu de PlanInfo.Exercices
Suite......... : BTY - 03/06 En Web Access l'exercice en cours n'était recherché que parmi
Suite......... : les exercices dont ex_etatcpta = OUV => OUV ou CPR
Suite......... : BTY - 04/06 FQ 17629 4 règles de calcul de la PMValue en cession au lieu d'une :
Suite......... : - CT Tout en court terme
Suite......... : - CT Tout en long terme
Suite......... : - NOR Règle normale (comme avant)
Suite......... : - RSD Règle normale en ignorant la durée (<=> NOR sans tenir compte de l'âge)
Suite......... : - MBO - fq 18343 - 08/06/2006 : pb calcul tableau prévisionnel si deb amort > N+1
Suite......... : - BTY - FQ 18370 - 06/06 Optimiser ChargeImmo en connaissant le order by du select
Suite......... : - MBO - FQ 18390 - dans le cas d'un dérog nul il faut prendre la dot éco +
                                    cas de la reprise dérogatoire jamais pris en compte dans
                                    les états prévisionnels
Suite......... : - MBO - FQ 15274 - 07/06 gestion plan variable avec antérieurs
Suite......... : - MBO - modif pour conseil de compilation - 24.08.06
Suite......... : - MVG 08/2006 Gestion des DPI
Suite......... : - MVG 31/08/2006 FQ18720 et FQ18720
Suite......... : - MBO 09/2006    - Gestion prime d'équipement
Suite......... : - MBO 09/2006    - Gestion Subvention d'équipement
Suite......... : - MBO 11/2006    - les bases éco et fiscales étaient obtenues par calcul
                                    maintenant on utilise les zones stockées dans tplanamort
Suite......... : - MBO 08/03/2008 - FQ 17512 - grosse modif pour calcul derogatoire et réintégration
Suite ........ : - BTY 07/07 FQ 21169 Revenir en arrière sur le tri optimisé de FQ 18370 car les infos immolog et immoamor de certaines
Suite ........ :   immos étaient oubliées d'où montants calculés à zéro (cf client SEGEC)
Suite ........ : - MBO 31/08/2007 - FQ 21309 - calcul d'une QP même si pas de plafond fiscal
Suite ........ : - MBO 10/09/2007 - FQ 21405 - cession  sur exercice postérieur à l'exercice de fin d'amortissement
Suite......... : - MBO 20/09/2007 - FQ 21424 - nvelle F° = GetPreviFiscalReel : prévision fiscale = fiscal réellement pratiqué
Suite......... : - MBO 05/10/2007 - FQ 19172 - nvelle F° = GetQteCedee = récupère la quantité pour impres° avt date cession (ilv)
Suite......... : - MBO 29/10/2007 - FQ 21754 - calcul avec amortisst ou non du jour de cession (Param supplémentaires en entrée de calcul)
Suite......... : - BTY 13/11/07 Fonctions d'édition de IFR immos issues du passage forfait à réel (Agricoles)
Mots clefs ... :
*****************************************************************}
unit ImPlanInfo;

interface

uses sYsuTILS, ParamSoc,
  HEnt1,
  HCtrls,
  UTob,
  ImPlan,
  ImPlanMeth
  {$IFDEF SERIE1}
  , S1Util  //YCP 25/08/05
  {$ELSE}
  , Ent1
  {$ENDIF}
{$IFDEF eAGLClient}
{$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
{$ENDIF eAGLClient}
  ;

const
  INDICE_LOG = 0;
  INDICE_AMORT = 1;
  TAILLE_CACHE = 25;

type
  TPlanInfo = class
    Plan: TPlanAmort;
    fbSortie: boolean;
    fbEclate: boolean;
    bExcep: boolean;
    DotationEco,
    DotationFisc,
    CumulAntEco,
    CumulAntFisc,
    CumulEco,
    CumulFisc,
    VNCEco,
    VNCFisc,
    CessionEco,
    CessionFisc,
    CumulCessionEco,
    CumulCessionFisc,
    BaseEco: double;
    BaseFisc: double;     // ajout mbo 15280
    ExcepEco: double;
    dTotExcEco : double;     // ajout mbo 16.11.05 et modif fq 17013
    dTotExcFisc : double;    // ajout mbo 30/11/05 - fq 17013
    DateRef: TDateTime;
    ValeurResiduelle: double;
    TVAAReverser: double;
    TotalPrixVente: double;
    NonDeductible: double;
    Reintegration: double;  // lié au plafond fiscal
    QuotePart: double;
    PValue: double;
    // BTY 04/06 FQ 17629 Règle de calcul de la PMValue en cession
    RegleCession : string;
    // BTY FQ 18370 Optimiser ChargeImmo en connaissant le order by du select
    SQLWhere : string;
    SQLOrderBy : string;

    fbPrime : boolean;   // Ajout mbo - présence prime d'équipement
    BasePri : double;
    DotationPri : double;
    CumulAntPri : double;
    CumulPri : double;
    VNCPri : double;
    CessionPri : double;
    CumulCessionPri : double;

    SBV : boolean;   // Ajout mbo - présence subvention
    BaseSbv : double;
    DotationSbv : double;
    CumulAntSbv : double;
    CumulSbv : double;
    VNCSbv : double;
    CessionSbv : double;
    CumulCessionSbv : double;

    DotationDR : double;     // ajout pour FQ 17512 Chantier fiscal
    DotationFEC: double;
    CumulAntDR: double;
    CumulAntFEC: double;
    CumulDR: double;
    CumulFEC: double;
    CumulCessionDR: double;
    CumulCessionFEC: double;
    GestionFiscale : boolean;

    constructor Create(CodeImmo: string);
    destructor Destroy; override;
  private
    //procedure Copie (Source : TPlanInfo) ;
    fTobEnCours: TOB; // Immo courante
    fTobImmo: TOB;
    fDateCalcul: TDateTime; // Mémorise la dernière date de calcul
    fTypeSortie: string;
    // Mémorise le type de sortie à la dernière date de calcul
    fDateSortie: TDateTime;
    fValAchatEclate: double;
    fTVARecuperable : double;   // ATTENTION : renseigné uniquement si sortie
    fTVARecuperee : double;     // ATTENTION : renseigné uniquement si sortie
    fValCedee : double;         // ATTENTION : renseigné uniquement si sortie fq 17289
    function GetPreviMeth(Indice: integer; MethodeAmort: TMethodeAmort): double;
    procedure AddTOBImmo(var OBDetail: TOB; CodeImmo: string);

    // mbo mise en public : function GetInfoCession(DateRef: TDateTime; var TVAAReverser, TotalPrixVente
    //  (*, dVNC*): double): boolean;
    procedure ChargeDetailImmo(T: TOB; Q: TQuery; iType: integer);
    function CourtTerme: boolean;
    procedure GetRepriseSuiteEclatement(DateCalcul: TDateTime; var RepriseEco:
      double; var RepriseFiscale: double);
    // FQ 17512
    procedure GetRepriseDREclatement(DateCalcul: TDateTime; var RepriseDR:
      double; var RepriseFEC: double);
    procedure GetReprisePriEclatement(DateCalcul: TDateTime; var ReprisePri:double);
    procedure GetRepriseSbvEclatement(DateCalcul: TDateTime; var RepriseSbv:double);

    // déplacement en public 20.09.07 function DateDebutDotation(DateRef: TDateTime): TDateTime;
    function DateFinDotation(DateRef: TDateTime): TDateTime;
    function EvalueDateExercice(DateRef: TDateTime; bFin: boolean): TDateTime;
    procedure CalculDerogatoire(DotEco, DotFisc, AE, AF, CE, CF, CumulDerog, CumulReint: double;
                                               var DotDerog, DotReint : double);
    function GetDotationPratiquee: double;
    function GetCumulAntPratique: double;
    function GetCessionFiscPratiquee: double;
  public
{$IFDEF EAGLSERVER}
    // BTY 02/06 FQ 17547  Montants faux en Web Access
    // => Démarrer à l'indice 0 pour pouvoir passer le tableau en paramètre
    Exercices: array[0..20] of TExoDate;  //array[1..20]
    // BTY 02/06 Pour les éditions avec TVA
    TobTva: TOB;
    // BTY 03/06 Pour les éditions si besoin du dégressif
    CoeffDegressif : TOB;
{$ENDIF}
    // ajout des 2 derniers param pour calcul avec ou sans prise en compte du jour de cession
    procedure Calcul(DateCalcul: TDateTime; AvecSortie: boolean; CalculPourCession : boolean; AmortJsortie : string = '');
    procedure UpdateTOB(CodeImmo: string);
    procedure ChargeImmo(CodeImmo: string);
    function GetPreviEco(Indice: integer): double;
    function GetPreviFisc(Indice: integer): double;
    function GetVoCedee(DateCalcul: TDateTime): double;

    // fq 19172
    function GetQteCedee(DateCalcul: TDateTime): double;

    function GetCumulAntEco(DateCalcul: TDateTime): double;
    function GetCumulAntFisc(DateCalcul: TDateTime): double;
    function GetCumulAntDR(DateCalcul: TDateTime; Reint :boolean = false): double;  // fq 17512
    //function GetCumulAntPRI(DateCalcul: TDateTime): double;
    function GetTypeSortie(DateCalcul: TDateTime): string;
    function GetValeurAchat(DateCalcul: TDateTime; bSortie: boolean): double;
    function GetPVCT: double;
    function GetPVLT: double;
    function GetDerogatoire : double;
    function GetReintegration : double;
    function GetCumulAntDerogatoire(DateRef : TDateTime) : double;   // les antérieurs
    function GetCumulAntFEC(DateRef : TDateTime): double;            // les antérieurs
    function GetCumulDerogatoireCede : double;    // antérieurs + dotation de l'exercice cédé
    function GetCumulFECcede : double;            // antérieurs + dotation de l'exercice cédé
    function GetPreviFiscalPratique (Indice: integer):double;
    function GetDotDeprec(DateCalcul: TDateTime): double;  // mbo 18.11.05

    function GetInfoCession(DateRef: TDateTime; var TVAAReverser, TotalPrixVente
      (*, dVNC*): double): boolean;

    function DateDebutDotation(DateRef: TDateTime): TDateTime;
    function GetPreviFiscalReel (Indice: integer):double; // FQ 21424 - mbo 20.09.07

    //BTY 13/11/07 Fonctions d'édition de IFR immos issues du passage forfait à réel (Agricoles)
    function GetVNCIFR(DateCalcul: TDateTime; bSortie: boolean): double;
    function GetDureeIFR(DateCalcul: TDateTime): double;

{$IFDEF EAGLSERVER}
    procedure SetExercicesServer(Exo: array of TExoDate);
    function SuivantServer: TExoDate;
    function EncoursServer: TExoDate;
    procedure UpdateInfo(CodeImmo: string; LaDate: TDateTime);
    // BTY 02/06 Pour les éditions Web Access avec TVA
    procedure SetTauxTVAServer(TobTvaWA:TOB);
    // BTY 03/06 Pour les éditions Web Access si besoin des coeff dégressifs
    procedure SetCoeffDegressifServer(CoeffDegressifWA:TOB);
{$ENDIF}
    published
      property Derogatoire : double read GetDerogatoire;
      property ReintFisc : double read GetReintegration;
      property DotationPratiquee : double read GetDotationPratiquee;
      property CumulAntPratique : double read GetCumulAntPratique;
      property CessionFiscPratiquee : double read GetCessionFiscPratiquee;
  end;

{$IFDEF EAGLSERVER}
function CQuelExerciceServer(Exercices: array of TExoDate; Date: TDateTime; var
  Exo: TExoDate): boolean;
{$ENDIF}

implementation
uses
{$IFDEF SERIE1}
{$ELSE}
  uLibExercice,
{$ENDIF}
  ImOuPlan,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ImEnt;

constructor TPlanInfo.Create(CodeImmo: string);
var
  Q: TQuery;
begin
  fDateCalcul := iDate1900;
  fTypeSortie := '';
  fTobImmo := TOB.Create('', nil, -1);
  fTobEnCours := nil;
  Plan := TPlanAmort.Create(True);
  if (CodeImmo <> '') then
  begin
    Q := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + CodeImmo + '"', True);
    // mbo fq 17054 en attente
    //if (Q.FindField('I_NATUREIMMO').AsString = 'CB') then
    //   fbSortie := (Q.FindField('I_OPELEVEEOPTION').AsString = 'X')
    //else
      fbSortie := (Q.FindField('I_OPECESSION').AsString = 'X');

    fbEclate := (Q.FindField('I_OPEECLATEMENT').AsString = 'X');
    bExcep := (Q.FindField('I_OPECHANGEPLAN').AsString = 'X');

    //fq 17512 chantier fiscal
    GestionFiscale := (Q.FindField('I_NONDED').AsString = 'X');

    fbPrime :=(Q.FindField('I_SBVPRI').AsFloat <> 0.00) or (Q.FindField('I_SBVPRIC').AsFloat <> 0.00) ;
    if not (fbsortie) then
       BasePri := (Q.FindField('I_SBVPRI').AsFloat)/2
    else
       BasePri := (Q.FindField('I_SBVPRIC').AsFloat)/2;

    SBV :=(Q.FindField('I_SBVMT').AsFloat <> 0.00) or (Q.FindField('I_SBVMTC').AsFloat <> 0.00);
    if (Q.FindField('I_SBVMT').AsFloat <> 0.00) then
       BaseSbv := (Q.FindField('I_SBVMT').AsFloat)
    else
       BaseSbv := (Q.FindField('I_SBVMTC').AsFloat);

    if not Q.Eof then
      Plan.Charge(Q);
    Ferme(Q);
  end;
end;

destructor TPlanInfo.Destroy;
begin
  fTobImmo.Free;
  if (Plan <> nil) then
    Plan.free;
// BTY 02/06
{$IFDEF EAGLSERVER}
  if (TobTva <> nil) then
     TobTva.free;
  // BTY 03/06
  if (CoeffDegressif <> nil) then
     CoeffDegressif.free;
{$ENDIF}
  inherited Destroy;
end;

//mbo - fq 21754 - 29/10/07 - nouveaux paramètres : CalculPourCession : est à true uniquement en appel par Imsortie et AmRemplacement
procedure TPlanInfo.Calcul(DateCalcul: TDateTime; AvecSortie: boolean; CalculPourCession : boolean; AmortJsortie : string = '');
var
  CumAntEco, CumAntFisc: double;
  Prorata, DotationEcoExo, DotationFiscExo,
    dDotEco, dCesEco, dExcEco, dExcDot, dDotFisc, dCesFisc, dExcFisc: double;
  FbSortieTotale: boolean;
  RepEco, RepFisc: double;
  DateOpeRevisionNam: TDateTime;
  AvecCloture : boolean;
  DotationPriExo, dDotPri,DCesPri :double;
  RepPri: double;
  DotationSbvExo, dDotSbv, DCesSbv : double;
  RepSbv: double;
  TraiterReprisePrem : boolean;
  CumAntDR, CumAntFEC :double;
  RepDR, RepFEC : double;
  DotationSeule : double;
  DotationZero, MoinsUnJour : boolean;
begin

  DateRef := DateCalcul;
  DotationEco := 0.00;
  CumulAntEco := 0.00;
  CumulEco := 0;
  VNCEco := 0;
  CessionEco := 0;
  CumulCessionEco := 0;
  ExcepEco := 0;
  DotationFisc := 0.00;
  CumulAntFisc := 0.00;
  CumulFisc := 0;
  VNCFisc := 0;
  CessionFisc := 0;
  CumulCessionFisc := 0;
  BaseEco := 0;
  BaseFisc:=0;
  dDotEco := 0;
  dCesEco := 0;
  dExcEco := 0;
  dExcDot := 0;
  dDotFisc := 0;
  dCesFisc := 0;
  dExcFisc := 0;
  TVAAReverser := 0;
  TotalPrixVente := 0;
  NonDeductible := 0;
  Reintegration := 0;
  QuotePart := 0;
  FbSortieTotale := false;
  fDateSortie := 2;
  DotationFiscexo := 0; // MVG 31/08/2006 FQ 18720 et FQ18721

  dDotPri := 0;
  dCesPri := 0;
  CumulPri := 0;
  CumulAntPri := 0;
  RepPri :=0;
  CessionPri := 0;
  CumulCessionPri :=0;
  dotationPri := 0;

  dDotSbv := 0;
  dCesSbv := 0;
  CumulSbv := 0;
  CumulAntSbv := 0;
  RepSbv :=0;
  CessionSbv := 0;
  CumulCessionSbv :=0;
  TraiterReprisePrem := false;

  RepDR := 0;
  RepFEC := 0;

  if fbSortie then
    FbSortieTotale := GetInfoCession(DateRef, TVAAReverser, TotalPrixVente (*,dVNC*));

  if fTobEnCours = nil then
    //    Plan.RecupereSuivantDate(iDate1900)
    Plan.RecupereSuivantDate(DateCalcul)
      //  else Plan.RecupereSuivantDate(iDate1900,fTobEnCours.Detail[INDICE_AMORT],fTobEnCours.Detail[INDICE_LOG]);
  else
    Plan.RecupereSuivantDate(DateCalcul, fTobEnCours.Detail[INDICE_AMORT],
      fTobEnCours.Detail[INDICE_LOG]);

  if FbSortieTotale then
     DateRef := DateFinDotation(DateRef);


  { Calcul des cumuls antérieurs }
  // BTY 03/06 Pb en édition Web Access avec VHIMMO
  {$IFDEF EAGLSERVER}
  AvecCloture := ( DateRef >= EncoursServer.Deb );
  {$ELSE}
  AvecCloture := ( DateRef >= VHImmo^.EnCours.Deb ); // CA - 17/09/2004 - FQ 13473 : calcul des cumuls antérieur pour des dates antérieures au début d'exercice.
  {$ENDIF}
  Plan.GetCumulsDotExercice(DateDebutDotation(DateRef), CumAntEco, CumAntFisc, AvecSortie, true, AvecCloture);

 if Plan.Fiscal then
     Plan.GetCumulsDotExerciceDR(DateDebutDotation(DateRef), CumAntDR, CumAntFEC, AvecSortie, true, AvecCloture);

 if fbPrime then
    Plan.GetCumulPri(DateDebutDotation(DateRef), CumulAntPri, AvecSortie, true, AvecCloture);


 if SBV then
 begin
   {$IFDEF EAGLSERVER}
   if (Plan.SBVDateSBV >= EncoursServer.Deb) and (Plan.SBVDateSBV >= EncoursServer.fin) and (Plan.AmortSBV.ReprisePrem) then
      TraiterReprisePrem := true;
  {$ELSE}
   if (Plan.SBVDateSBV >= VHImmo^.EnCours.Deb) and (Plan.SBVDateSBV >= VHImmo^.EnCours.Fin) and (Plan.AmortSBV.ReprisePrem) then
      TraiterReprisePrem := true;
  {$ENDIF}

   Plan.GetCumulSbv(DateDebutDotation(DateRef), CumulAntSbv, AvecSortie, true, AvecCloture);

 end;

  { Calcul des dotations }
  DotationEcoExo := Plan.GetDotationGlobale(1, DateFinDotation(DateRef),
                                              dDotEco, dCesEco, dExcEco);

  { On récupère les elements exceptionnels de l'exercice que l'on soustrait à la dotation de l'exercice }
  if bExcep then
    // modif mbo 24.10.05 dExcDot := Plan.GetExcepExercice(DateRef, DateOpeRevisionNam, True);
    dExcDot := Plan.GetExcepExercice(DateRef, DateOpeRevisionNam, True, false);

  (*  if fbSortie then  // Pour ne pas perturber les autres calculs (exceptionnel déjà pris en compte dans la sortie
    begin
      ExcepSortie := dExcDot;
      dExcDot := 0;
    end;*)
    { Si une révision NAM vers AM a eu lieu, on récupère la date d'opération pour les calculs }
  if DateOpeRevisionNam > iDate1900 then
  begin
    Plan.DateMiseEnService := DateOpeRevisionNam;
    // fq 17569 ajout pour nouvelle date deb amort
    Plan.DateDebEco := DateOpeRevisionNam;
  end;

  //new mbo 28.09.07
  DotationZero := false;
  MoinsUnJour := false;

  if (CalculPourCession) then
  begin
      if DateDebutDotation(DateRef) = DateRef then   //date de cession = 1er jour de l'exercice
      begin
         if AmortJsortie = 'NON' then    //jour de cession non pris en cpte ds l'amortissement
            DotationZero:= true
         else
            DotationZero:= false;
      end else
      begin
         if AmortJsortie = 'NON' then
            MoinsUnJour:= true        //jour de cession non pris en cpte ds l'amortissement
         else
            DotationZero:= false;
      end;
   end else
   begin
      DotationZero := false;
      MoinsUnJour:= false;
   end;
  // modif mbo appel des fonctions avec les dates datedebeco, datedbfis et non mise en service ou achat
  // 27.09 if ((Plan.AmortEco.Methode = 'LIN') or  (Plan.AmortEco.Methode = 'VAR')) then   // CA - 20/12/2004 - on gère le variable comme le linéaire
    //    DotationEco:=GetDotationLinAvant(VHImmo^.EnCours.Deb,VHImmo^.EnCours.Fin,Plan.DateMiseEnService,DateRef,Plan.AmortEco.DateFinAmort,DotationEcoExo,False)
  if (Plan.AmortEco.Methode = 'LIN') then
  begin
    //new
    if DotationZero then
       DotationSeule := 0

    else if MoinsUnJour then
       DotationSeule := GetDotationLinSortie(DateDebutDotation(DateRef),
                                             DateFinDotation(DateRef), Plan.DateDebEco, DateRef,
                                             Plan.AmortEco.DateFinAmort, (DotationEcoExo - dExcDot), False)
    else
       DotationSeule := GetDotationLinAvant(DateDebutDotation(DateRef),
                                            DateFinDotation(DateRef), Plan.DateDebEco, DateRef,
                                            Plan.AmortEco.DateFinAmort, (DotationEcoExo - dExcDot), False);
    {
    DotationEco := GetDotationLinAvant(DateDebutDotation(DateRef),
                                       DateFinDotation(DateRef), Plan.DateDebEco, DateRef,
                                       Plan.AmortEco.DateFinAmort, (DotationEcoExo - dExcDot), False)
                                       + dExcDot
    }
    DotationEco := DotationSeule + dExcDot;
  end else
  begin
    if (Plan.AmortEco.Methode = 'VAR') then
        DotationEco := GetDotationLinAvant(DateDebutDotation(DateRef),
                                       DateFinDotation(DateRef), Plan.DateDebEco, DateRef,
                                       Plan.AmortEco.DateFinAmort, (DotationEcoExo - dExcDot), False)
                                       + dExcDot


    else
    //    DotationEco:=GetDotationDegAvant(VHImmo^.EnCours.Deb,VHImmo^.EnCours.Fin,Plan.DateAchat,DateRef,DotationEcoExo,Prorata, Plan.Suramort);
        DotationEco := GetDotationDegAvant(DateDebutDotation(DateRef),
                                       DateFinDotation(DateRef), Plan.DateDebEco, DateRef,
                                       (DotationEcoExo - dExcDot), Prorata, Plan.Suramort)
                                       + dExcDot;
  end;
  //YCP 12/09/01 DotationEco:=DotationEco+dExcDot ;
  //YCP 12/09/01 if AvecSortie then DotationEco:=DotationEco+dCesEco ; // dot immo - elem exc exe + elem exe prorata
  // YCP 12/09/01 deb
  //YCP 12/09/01 fin
  if Plan.Fiscal then
  begin
    // MVG 31/08/2006 pour les FQ 18720 et FQ 18721
    //    Plan.GetDotationGlobale(2, DateFinDotation(DateRef), dDotFisc, dCesFisc,
    //                            dExcFisc); //EPZ 07/11/00
    //    DotationFiscExo := dDotFisc
    // Fin MVG 31/08/2006
    DotationFiscExo := Plan.GetDotationGlobale(2, DateFinDotation(DateRef), dDotFisc, dCesFisc,dExcFisc);

    if Plan.AmortFisc.Methode = 'LIN' then
       {mbo - fq 15280 - dot exceptionnelle ne doit pas être proratisée
        DotationFisc := GetDotationLinAvant(DateDebutDotation(DateRef),
        DateFinDotation(DateRef), Plan.DateMiseEnService, DateRef,
        Plan.AmortFisc.DateFinAmort, DotationFiscExo, false);}
    begin
        if DotationZero then
           DotationSeule := 0
        else if MoinsUnJour then
           DotationSeule := GetDotationLinSortie(DateDebutDotation(DateRef),
                                            DateFinDotation(DateRef), Plan.DateDebFis, DateRef,
                                            Plan.AmortFisc.DateFinAmort, DotationFiscExo-dExcDot, false)
        else
           DotationSeule := GetDotationLinAvant(DateDebutDotation(DateRef),
                                            DateFinDotation(DateRef), Plan.DateDebFis, DateRef,
                                            Plan.AmortFisc.DateFinAmort, DotationFiscExo-dExcDot, false);

        DotationFisc := DotationSeule + dExcDot;

      // Fiscal = Economique + dérogatoire signé
      // DotationFisc := DotationEco + Derogatoire;
    end else
    begin
      {mbo - fq 15280 on ne proratise l'élement exceptionnel
      DotationFisc := GetDotationDegAvant(DateDebutDotation(DateRef),
        DateFinDotation(DateRef), Plan.DateAchat, DateRef, DotationFiscExo,
        Prorata, Plan.SurAmort); }

        DotationFisc := GetDotationDegAvant(DateDebutDotation(DateRef),
                                            DateFinDotation(DateRef), Plan.DateDebFis, DateRef,
                                            (DotationFiscExo-dExcDot),Prorata, Plan.SurAmort)
                                            + dExcDot;
    end;

    if FbSortieTotale then
    begin
      if Plan.GetVNCCedee(Plan.AmortFisc, DateFinDotation(DateRef)) <= 0 then
          VNCFisc := 0
      else
          VNCFisc := Plan.GetVNCCedee(Plan.AmortFisc, DateFinDotation(DateRef)) -
                                      Plan.AmortFisc.RepriseCedee;
    end else
    begin
        VNCFisc := Plan.GetVNCAvecMethode(Plan.AmortFisc, DateRef - 1) -
                                          DotationFisc;
    end;

    VNCFisc := Arrondi(VNCFisc, V_PGI.OkDecV);
  end;

  if Plan.PRI then
  begin
    DotationPriExo := Plan.GetDotationGlobale(3, DateFinDotation(DateRef), dDotPri, dCesPri,dExcFisc);
    if Plan.AmortPri.Methode = 'LIN' then
        DotationPri := GetDotationLinAvant(DateDebutDotation(DateRef),
                                           DateFinDotation(DateRef), Plan.PriDateDeb, DateRef,
                                           Plan.AmortPri.DateFinAmort, DotationPriExo, false)
    else
        DotationPri := GetDotationDegAvant(DateDebutDotation(DateRef),
                                           DateFinDotation(DateRef), Plan.PriDateDeb, DateRef, DotationPriExo,
                                           Prorata, Plan.SurAmort);
  end;

  if Plan.SBV then
  begin
    DotationSbvExo := Plan.GetDotationGlobale(4, DateFinDotation(DateRef), dDotSbv, dCesSbv,dExcFisc);

    // Ds le cas ou la reprise est cumulée avec la première dotation, si la date de calcul est
    // différente de la date fin d'exercice, la partie reprise contenue dans la dotation ne
    // doit pas être proratisée ---> il faut donc l'enlever avant la proratisation et la remettre après

    if TraiterReprisePrem then
          DotationSbvExo := DotationSbvExo - Plan.AmortSBV.Reprise;


    if Plan.AmortSbv.Methode = 'LIN' then
        DotationSbv := GetDotationLinAvant(DateDebutDotation(DateRef),
                                           DateFinDotation(DateRef), Plan.SbvDateDeb, DateRef,
                                           Plan.AmortSbv.DateFinAmort, DotationSbvExo, false)
    else
        DotationSbv := GetDotationDegAvant(DateDebutDotation(DateRef),
                                           DateFinDotation(DateRef), Plan.SbvDateDeb, DateRef, DotationSbvExo,
                                           Prorata, Plan.SurAmort);


    if TraiterReprisePrem = true then
       DotationSbv := DotationSbv + Plan.AmortSBV.Reprise;

  end;

  if AvecSortie then
  begin
    CumulAntEco := CumAntEco + Plan.AmortEco.RepriseCedee;
    CumulAntFisc := CumAntFisc + Plan.AmortFisc.RepriseCedee;
    CumulAntDR := CumAntDr + Plan.AmortDerog.RepriseCedee;
    CumulAntFEC := CumAntFEC + Plan.AmortReint.RepriseCedee;
  end else
  begin
    CumulAntEco := CumAntEco;
    CumulAntFisc := CumAntFisc;
    CumulAntDR := CumAntDr;
    CumulAntFEC := CumAntFEC;
  end;
  { CA - 21/01/2004 - Gestion éclatement en cas de reprise }
  if fbEclate then
  begin
    GetRepriseSuiteEclatement(DateRef, RepEco, RepFisc);
    GetRepriseDREclatement(DateRef, RepDR, RepFEC);
    if Plan.AmortEco.Reprise <> 0 then
    begin
      CumulAntEco := CumulAntEco + RepEco;
      CumulAntFisc := CumulAntFisc + RepFisc;
      CumulAntDR := CumulAntDr + RepDr;
      CumulAntFEC := CumulAntFec + RepFec;
    end;

    GetReprisePriEclatement(DateRef, RepPri);
    if Plan.AmortPri.Reprise <> 0 then
        CumulAntPri := CumulAntPri + RepPri;

    // ajout mbo pour reprise sbv
    GetRepriseSbvEclatement(DateRef, RepSbv);
    //if (Plan.AmortSbv.Reprise <> 0) and (Plan.AmortSbv.ReprisePrem = false) then
    if (Plan.AmortSbv.Reprise <> 0) and (TraiterReprisePrem = false) then
        CumulAntSbv := CumulAntSbv + RepSbv;

  end;
  { FIN CA - 21/01/2004 - Gestion éclatement en cas de reprise }
  CumulAntEco := Arrondi(CumulAntEco, V_PGI.OkDecV);
  CumulAntFisc := Arrondi(CumulAntFisc, V_PGI.OkDecV);
  CumulEco := Arrondi(CumulAntEco + DotationEco, V_PGI.OkDecV);
  CumulFisc := Arrondi(CumulAntFisc + DotationFisc, V_PGI.OkDecV);

  CumulAntPri := Arrondi(CumulAntPri, V_PGI.OkDecV);
  CumulPri := Arrondi(CumulAntPri + DotationPri, V_PGI.OkDecV);
  CumulAntSbv := Arrondi(CumulAntSbv, V_PGI.OkDecV);
  CumulSbv := Arrondi(CumulAntSbv + DotationSbv, V_PGI.OkDecV);

  if Plan.AmortEco.Methode = 'NAM' then
  begin
    if fbSortie then
      VNCEco := Plan.ValeurAchat
    else
      VNCEco := Plan.AmortEco.Base;
  end else
  begin
    if FbSortieTotale then
    begin
      if Plan.GetVNCCedee(Plan.AmortEco, DateFinDotation(DateRef)) = 0 then
        VNCEco := 0
          //    else VNCEco := Plan.GetVNCCedee(Plan.AmortEco,VHImmo^.Encours.Fin) - Plan.AmortEco.RepriseCedee;
      // CA - 02/06/2003 - Prise en compte de l'exceptionnel dans la VNC cédée.
      else
        VNCEco := Plan.GetVNCCedee(Plan.AmortEco, DateFinDotation(DateRef)) -
                  Plan.AmortEco.RepriseCedee + fTVARecuperable-fTVARecuperee+TVAAReverser;
                  // Plan.AmortEco.RepriseCedee - dExcDot;  // FQ 12370
    end else
    begin
        VNCEco := Plan.GetVNCAvecMethode(Plan.AmortEco, DateRef - 1) - DotationEco;
        // fq 17289 - mbo 10.01.2006 fbsortie à true = cession dans l'exo
        //            vnc négative car getVNCavec methode renvoie 0 ds ce cas
        if (VNCEco < 0) AND (fbSortie = true) then
        begin
           // cas d'une cession avec édition avant date de cession
           // on reconstitue la valeur d'achat car le montantht est à 0

           // BTY 03/06 Pb en édition Web Access avec VHIMMO
          {$IFDEF EAGLSERVER}
           if GetInfoCession(EnCoursServer.fin , TVAAReverser, TotalPrixVente) then
          {$ELSE}
           if GetInfoCession(VHImmo^.EnCours.fin , TVAAReverser, TotalPrixVente) then
          {$ENDIF}

           begin
              // dans fValCedee on a le montant HT
              VNCEco:= fValCedee - fTvaRecuperable + fTvaRecuperee;
              // fq 21405 - 06.09.07 VNCEco := VNCEco - Plan.AmortEco.RepriseCedee - DotationEco;
              VNCEco := VNCEco - CumulEco;
           end;
        end else   // ajout mbo fq 21405 - 06.09.07
        begin
          // if (fbSortie = true) then
          //    VNCEco:= VNCEco - Plan.AmortEco.RepriseCedee - Plan.AmortEco.RepriseDep;
        end;
    end;
    VNCEco := Arrondi(VNCEco, V_PGI.OkDecV);
  end;

  CessionEco := Arrondi(dCesEco, V_PGI.OkDecV);
  CessionFisc:=Arrondi (dCesFisc,V_PGI.OkDecV);

  CessionPri:=Arrondi (dCesPri,V_PGI.OkDecV);
  CessionSbv:=Arrondi (dCesSbv,V_PGI.OkDecV);

  if FbSortieTotale then
  begin
     //if Plan.GetVNCCedee(Plan.AmortPri, DateFinDotation(DateRef)) = 0 then
     VNCPri := 0
     //else
     //   VNCPri := Plan.GetVNCCedee(Plan.AmortPri, DateFinDotation(DateRef)) -
     //                              Plan.AmortPri.RepriseCedee;
  end else
  begin
     VNCPri := Plan.GetVNCAvecMethode(Plan.AmortPri, DateRef - 1) - DotationPri;

     if (VNCPri < 0) AND (fbSortie = true) then
     begin
        // cas d'une cession avec édition avant date de cession
        VNCPri := BasePri + VNCPri - Plan.AmortPri.RepriseCedee;       //- Plan.AmortPri.RepriseCedee - DotationPri;
     end;
  end;
  VNCPri := Arrondi(VNCPri, V_PGI.OkDecV);

  // pour les subventions
  if FbSortieTotale then      // à fin d'exercice de cession la vnc est toujours à 0
    VNCSbv := 0.00
  else
  begin
     // Cette fonction renvoie le montant des antérieurs (en négatif si cession car la base est à 0)
     VNCSbv := Plan.GetVNCAvecMethode(Plan.AmortSbv, DateRef - 1) - DotationSbv;

     if (VNCSbv < 0) AND (fbSortie = true) then // cession ds exercice avec édition avant cession

     begin
        if Plan.AmortSBV.ReprisePrem then
           VNCSbv := BaseSbv + VNCSbv
        else
           VNCSbv := BaseSbv + VNCSbv - Plan.AmortSBV.RepriseCedee;
     end;
  end;
  VNCSbv := Arrondi(VNCSbv, V_PGI.OkDecV);

  (*  if fbSortie then ExcepEco := ExcepSortie
    else *)
  ExcepEco := Arrondi(dExcDot, V_PGI.OkDecV);

  {ajout mbo 16.11.05 alimentation du total exceptionnel + dépréciation actif
   dans tplaninfo pour les éditions}
  dTotExcEco := Plan.GetExcepExercice(DateRef, DateOpeRevisionNam, True, True);
  dTotExcEco := Arrondi(dTotExcEco, V_PGI.OkDecV);

  // ajout mbo 30/11/05 - fq 17013
  // fq 17476 impact depreciation sur fiscal sauf si plan éco = VAR
  // TotexcDeprec := Plan.GetExcepExercice(DateRef, DateOpeRevisionNam, True, False);
  if Plan.AmortEco.Methode = 'VAR' then
     dTotExcFisc := Plan.GetExcepExercice(DateRef, DateOpeRevisionNam, True, False)
  else
     dTotExcFisc := Plan.GetExcepExercice(DateRef, DateOpeRevisionNam, True, True);

  dTotExcFisc := Arrondi(dTotExcFisc, V_PGI.OkDecV);

  Plan.GetCumulDotationCession(CumulCessionEco, CumulCessionFisc, DateDebutDotation(DateRef));

  Plan.GetCumulDotCessionPRI(CumulCessionPri,DateDebutDotation(DateRef));
  Plan.GetCumulDotCessionSBV(CumulCessionSbv,DateDebutDotation(DateRef));

  ValeurResiduelle := Plan.GetBaseExercice(DateRef, Plan.AmortEco) - DotationEco;

// MVG 08/2006 Gestion des DPI
// MVG 31/08/2006 FQ18720 et FQ18720

 QuotePart := 0.0;
 If (Plan.DPIAffectee<>0.0) then
 begin
   QuotePart := 0.0;
   if Plan.Fiscal then
      Reintegration := DotationFiscExo*(Plan.CoeffReintegration)
   else
      Reintegration :=DotationEcoExo*(Plan.CoeffReintegration);
 end;

 If (Plan.DPIAffectee=0.0) then
 begin  // ajout du begin - mbo - 29.09.06
     if Plan.Fiscal then
     begin
        if Plan.CoeffReintegration > 0.0 then
        begin
           Reintegration := DotationFiscExo*(1 - Plan.CoeffReintegration);
           QuotePart := DotationFiscExo*Plan.CoeffReintegration*Plan.TauxQuotePart/100;
        end else  // fq 21309 ajout du calcul QP même si pas de plafond fiscal
           QuotePart := DotationFiscExo*Plan.TauxQuotePart/100;
     end else
     begin
        if Plan.CoeffReintegration > 0.0 then
        begin
           Reintegration := DotationEcoExo*(1 - Plan.CoeffReintegration);
           QuotePart := DotationEcoExo*Plan.CoeffReintegration*Plan.TauxQuotePart/100;
        end else   // fq 21309 ajout du calcul QP même si pas de plafond fiscal
           QuotePart := DotationEcoExo*Plan.TauxQuotePart/100;

     end;
 end; // ajout du end - mbo - 29.09.06

  {MVG 31/08/2006 FQ18720 et FQ18721
  { CA - 04/06/2003 - calcul de la valeur de l'amortissement non déductible }
  {  if (SommeND <> 0) or (Plan.TauxQuotePart <> 0) then
  begin
    // Calcul de la réintégration fiscale
    if ((SommeND <> 0) and (Plan.AmortEco.Base <> 0)) then
        Reintegration := (1 - (SommeND / Plan.AmortEco.Base)) *
          DotationEcoExo
          // CA - 20/01/2004 - FQ 13217 - Calcul de la réintégration si sortie.
      else if ((SommeND <> 0) and (Plan.ValeurAchat <> 0)) then
        Reintegration := (1 - (SommeND / Plan.ValeurAchat)) *
          DotationEcoExo;
      NonDeductible := Reintegration + ((Plan.TauxQuotePart / 100) *
        (DotationEcoExo - Reintegration));
      if ((SommeND <> 0) and (Plan.AmortEco.Base <> 0)) then
        QuotePart := (DotationEcoExo / Plan.AmortEco.Base) * SommeND
          * (Plan.TauxQuotePart / 100)
      else if ((SommeND <> 0) and (Plan.ValeurAchat <> 0)) then
        QuotePart := (DotationEcoExo / Plan.ValeurAchat) * SommeND *
         (Plan.TauxQuotePart / 100);
    end;}

  { CA - 10/06/2003 - Calcul de la base économique }
  if Plan.AmortEco.Methode = 'NAM' then
      BaseEco := Plan.AmortEco.Base // FQ 14093
  else
  begin
    if fbSortieTotale then
    begin
      BaseEco := Plan.AmortEco.BaseDebExoCes[0];
      // ajout mbo fq 15280
      BaseFisc:= Plan.AmortFisc.BaseDebExoCes[0];
    end else
    begin
      // modif mbo 14.11.06 - pour impression à une date antérieure à une modif de base
      {BaseEco := Plan.AmortEco.BaseDebutExo[0] + Plan.AmortEco.Reprise +
        Plan.AmortEco.RepriseCedee;
      // ajout mbo fq 15280
      BaseFisc := Plan.AmortFisc.BaseDebutExo[0] + Plan.AmortFisc.Reprise +
        Plan.AmortFisc.RepriseCedee;
      }
      BaseEco := Plan.AmortEco.Base;
      BaseFisc := Plan.AmortFisc.Base;

    end;
  end;
end;
//---------------------------------------------

function TPlanInfo.GetPreviEco(Indice: integer): double;
begin
  result := GetPreviMeth(indice, Plan.AmortEco);
end;
//=======================================================

function TPlanInfo.GetPreviFisc(Indice: integer): double;
begin
    result := GetPreviFiscalPratique (indice);
end;

//=========================================
function TPlanInfo.GetPreviMeth(Indice: integer; MethodeAmort: TMethodeAmort):
  double;
var
  i: integer;
  dDebut, dFin: TDateTime;
begin
  i := 0;
  CalculDateDebutFinExByPos(Indice, dDebut, dFin);
  while ((Plan.TableauDate[i] <> iDate1900) and (Plan.TableauDate[i] <
    DateFinDotation(DateRef))) do
    Inc(i);

  // fq 18343 - mbo - 08.06.2006  - ajout du test ds le cas ou immo amortie sur exo > N-1
  if Plan.TableauDate[i] > dFin then
    result := 0
  else
  begin
    while ((Plan.TableauDate[i] <> iDate1900) and (Plan.TableauDate[i] < dFin)) do
      Inc(i);
     result := MethodeAmort.TableauDot[i];
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/01/2005
Modifié le ... : 20/09/2007
Description .. : Calcul l'amortissement fiscal en prenant en compte dot éco
Suite ........ : + dérogatoire
Mots clefs ... : 
*****************************************************************}
function TPlanInfo.GetPreviFiscalPratique (Indice: integer):  double;
var
  i: integer;
  dDebut, dFin: TDateTime;
begin
  i := 0;
  CalculDateDebutFinExByPos(Indice, dDebut, dFin);
  while ((Plan.TableauDate[i] <> iDate1900) and (Plan.TableauDate[i] <
          DateFinDotation(DateRef))) do
    Inc(i);

  while ((Plan.TableauDate[i] <> iDate1900) and (Plan.TableauDate[i] < dFin)) do
      Inc(i);

  Result := Plan.AmortEco.TableauDot[i] + Plan.AmortDerog.TableauDot[i];
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 20/09/2007
Modifié le ... :   /  /    
Description .. : Prévisions fiscales : prise en compte de la dotation fiscale 
Suite ........ : réellement pratiquée
Mots clefs ... : 
*****************************************************************}
function TPlanInfo.GetPreviFiscalReel (Indice: integer):  double;
var
  i: integer;
  dDebut, dFin: TDateTime;
begin
  i := 0;
  CalculDateDebutFinExByPos(Indice, dDebut, dFin);
  while ((Plan.TableauDate[i] <> iDate1900) and (Plan.TableauDate[i] <
          DateFinDotation(DateRef))) do
    Inc(i);

  while ((Plan.TableauDate[i] <> iDate1900) and (Plan.TableauDate[i] < dFin)) do
      Inc(i);

  Result := Plan.AmortFisc.TableauDot[i];
end;

//=====================================================================

function TPlanInfo.GetInfoCession(DateRef: TDateTime; var TVAAReverser,
  TotalPrixVente (*,dVNC*): double): boolean;
var // Q : TQuery;
  TypOp: string;
  i: integer;
begin
  TVAAReverser := 0;
  fTVARecuperable := 0;
  fTVARecuperee := 0;
  TotalPrixVente := 0; (*dVNC := 0;*)
  result := False;

  for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
  begin
    if fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_DATEOP') <
      DateDebutDotation(DateRef) then
      continue;
    if fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_DATEOP') > DateRef
      then
      break;

    TypOp := fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_TYPEOP');
    if (TypOp = 'CES') or (TypOp = 'CEP') then
    begin
      TVAAReverser := TVAAReverser +
        fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_TVAAREVERSER');
      TotalPrixVente := TotalPrixVente +
        fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_MONTANTCES');
      fDateSortie :=
        fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_DATEOP');
      PValue := fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_PVALUE');
      fTVARecuperable := fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_TVARECUPERABLE');
      fTVARecuperee := fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_TVARECUPEREE');

      // ajout mbo - fq 17289
      fValCedee := fTobEnCours.Detail[INDICE_LOG].Detail[i].GetValue('IL_VOCEDEE');
    end;
    if not result then
      result := (TypOp = 'CES');
  end;
end;

//=========================================================================

procedure TPlanInfo.UpdateTob(CodeImmo: string);
var
  T, T2: TOB;
begin
  fDateCalcul := iDate1900;
  fTypeSortie := '';
  if (fTobEncours <> nil) then
    fTobEnCours.ClearDetail;
  //YCP 31-07-01 if VHImmo^.OBImmo<>nil then T:=VHImmo^.OBImmo.FindFirst(['I_IMMO'],[CodeImmo],False);
{$IFNDEF IMP}
  //YCP 31-07-01  if T=nil then T:=AddTobImmo(CodeImmo);
  T := TOB.Create('IMMO', nil, -1);
  AddTobImmo(T, CodeImmo);
{$ENDIF}
  if T.Detail.Count <> 0 then
  begin
    T2 := T.Detail[0];
    Plan.Init;
    fbSortie := (T2.GetValue('I_OPECESSION') = 'X');
    bExcep := (T2.GetValue('I_OPECHANGEPLAN') = 'X');
    fbEclate := (T2.GetValue('I_OPEECLATEMENT') = 'X');
    fbPrime := ((T2.GetValue('I_SBVPRI') <> 0.00) or (T2.GetValue('I_SBVPRIC')<> 0.00));
    SBV := ((T2.GetValue('I_SBVMT') <> 0.00) or (T2.GetValue('I_SBVMTC') <> 0.00));

    Plan.ChargeTOB(T2);
  end;
  T.Free;
end;

procedure TPlanInfo.AddTOBImmo(var OBDetail: Tob; CodeImmo: string);
var
  Q: TQuery;
  // OBAmor: TOB;
begin
  if OBDetail = nil then
    exit;
  //  Q := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+CodeImmo+'"',True);
  Q := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO>="' + CodeImmo + '"', True,
    TAILLE_CACHE);
  //YCP 31-07-01if VHImmo^.OBImmo=nil then VHImmo^.OBImmo:=Tob.Create('',nil,-1) ;
  //YCP 31-07-01VHImmo^.OBImmo.LoadDetailDB('IMMO','','',Q,True,True) ;
  OBDetail.LoadDetailDB('IMMO', '', '', Q, False, True);
  Ferme(Q);
  //YCP 31-07-01 OBDetail:=VHImmo^.OBImmo.FindFirst(['I_IMMO'],[CodeImmo],False);
  // CA - 10/10/2002 - Inutile, jamais utilisé.
(*  if OBDetail <> nil then
  begin
    Q:=OpenSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+CodeImmo+'" ORDER BY IA_NUMEROSEQ DESC, IA_DATE',True);
    OBAmor:=TOB.Create('',OBDetail,-1);
    OBAmor.LoadDetailDB('IMMOAMOR','','',Q,False,False);
    Ferme (Q);
  end;*)
  //YCP 31-07-01 result:=OBDetail;
end;

{procedure TPlanInfo.Copie(Source: TPlanInfo);
begin
  Plan.Copie(Source.Plan);
  bSortie         :=Source.bSortie;
  bExcep          :=Source.bExcep;
  DotationEco     :=Source.DotationEco;
  DotationFisc    :=Source.DotationFisc;
  CumulAntEco     :=Source.CumulAntEco;
  CumulAntFisc    :=Source.CumulAntFisc;
  CumulEco        :=Source.CumulEco;
  CumulFisc       :=Source.CumulFisc;
  VNCEco          :=Source.VNCEco;
  VNCFisc         :=Source.VNCFisc;
  CessionEco      :=Source.CessionEco;
  CessionFisc     :=Source.CessionFisc;
  ExcepEco        :=Source.ExcepEco;
  CumulCessionEco :=Source.CumulCessionEco;
  CumulCessionFisc:=Source.CumulCessionFisc;
  DateRef         :=Source.DateRef;
  ValeurResiduelle:=Source.ValeurResiduelle;
  TVAAReverser    :=Source.TVAAReverser;
  TotalPrixVente  :=Source.TotalPrixVente;
end; }

function TPlanInfo.GetVoCedee(DateCalcul: TDateTime): double;
var
  i: integer;
  T: TOB;
begin
  Result := 0;

  // pour fq 19172
  if DateCalcul = iDate1900 then
  begin
      for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
      begin
        T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
        if T.GetValue('IL_TYPEOP') = 'CES' then
        begin
          Result := T.GetValue('IL_VOCEDEE');
          break;
        end;
      end;

  end else
  begin
  for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
  begin
    T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
    if T.GetValue('IL_DATEOP') > DateCalcul then
      break;
    if T.GetValue('IL_TYPEOP') = 'CES' then
    begin
      Result := T.GetValue('IL_VOCEDEE');
      break;
    end;
  end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 05/10/2007
Modifié le ... :   /  /
Description .. : fonction permettant de récupérer la quantité.
Suite ........ : sert pour impression demandée avant date de cession d'une
Suite ........ : immo
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetQteCedee(DateCalcul: TDateTime): double;
var
  i: integer;
  T: TOB;
begin
  Result := 0;

  // pour fq 19172
  if DateCalcul = iDate1900 then
  begin
      for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
      begin
        T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
        if T.GetValue('IL_TYPEOP') = 'CES' then
        begin
          Result := T.GetValue('IL_QTECEDEE');
          break;
        end;
      end;

  end else
  begin
  for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
  begin
    T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
    if T.GetValue('IL_DATEOP') > DateCalcul then
      break;
    if T.GetValue('IL_TYPEOP') = 'CES' then
    begin
      Result := T.GetValue('IL_QTECEDEE');
      break;
    end;
  end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/04/2004
Modifié le ... :   /  /
Description .. : Calcul du cumul antérieur économique
Mots clefs ... :
*****************************************************************}

function TPlanInfo.GetCumulAntEco(DateCalcul: TDateTime): double;
var
  i: integer;
  T: TOB;
begin
  Result := 0;
  if fbSortie then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      // modif mbo 15.05.07
      //if T.GetValue('IL_DATEOP') > DateCalcul then
      //  break;

      if T.GetValue('IL_TYPEOP') = 'CES' then
        Result := Result + T.GetValue('IL_CUMANTCESECO');
    end;
  end
  else
    Result := CumulAntEco;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/04/2004
Modifié le ... :   /  /
Description .. : Calcul du cumul antérieur fiscal
Mots clefs ... :
*****************************************************************}
//
function TPlanInfo.GetCumulAntFisc(DateCalcul: TDateTime): double;
var
  i: integer;
  T: TOB;
begin
  Result := 0;
  if fbSortie then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      // modif mbo 15.05.07
      //if T.GetValue('IL_DATEOP') > DateCalcul then
      //  break;

      if T.GetValue('IL_TYPEOP') = 'CES' then
        Result := Result + T.GetValue('IL_CUMANTCESFIS');
    end;
  end
  else
    Result := CumulAntFisc;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : recup des antérieurs dérogatoires ou anterieurs réintégration
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetCumulAntDR(DateCalcul: TDateTime; reint : boolean): double;
var
  i: integer;
  T: TOB;
begin
  Result := 0;
  if fbSortie then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      if T.GetValue('IL_DATEOP') > DateCalcul then
        break;
      if T.GetValue('IL_TYPEOP') = 'CES' then
      begin
        if reint = false then
           Result := Result + T.GetValue('IL_DOUBLE1')
        else
           Result := Result + T.GetValue('IL_DOUBLE2');
      end;
    end;
  end
  else
    Result := CumulAntFisc;
end;


function TPlanInfo.GetTypeSortie(DateCalcul: TDateTime): string;
var
  i: integer;
  T: TOB;
begin
  Result := '';
  if DateCalcul = fDateCalcul then
    Result := fTypeSortie
  else
  begin
    for i := fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 downto 0 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      if T.GetValue('IL_DATEOP') > DateCalcul then
        continue;
      if Copy(T.GetValue('IL_TYPEOP'), 1, 2) = 'CE' then
      begin
        fTypeSortie := T.GetValue('IL_TYPEOP');
        fDateCalcul := DateCalcul;
        result := fTypeSortie;
        break;
      end;
    end;
  end;
end;

procedure TPlanInfo.ChargeDetailImmo(T: TOB; Q: TQuery; iType: integer);
var
  stLastCode, stNewCode: string;
  TLog: TOB;
  i, iMem: integer;
  stTable, stImmo: string;
begin
  stLastCode := '';
  iMem := 0;
  if iType = INDICE_LOG then
  begin
    stTable := 'IMMOLOG';
    stImmo := 'IL_IMMO';
  end
  else if iType = INDICE_AMORT then
  begin
    stTable := 'IMMOAMOR';
    stImmo := 'IA_IMMO';
  end;
  while not Q.Eof do
  begin
    stNewCode := Q.FindField(stImmo).AsString;
    if stLastCode <> stNewCode then
    begin
      for i := iMem to T.Detail.Count - 1 do
      begin
        if T.Detail[i].GetValue('I_IMMO') = stNewCode then
        begin
          TLog := TOB.Create(stTable, T.Detail[i].Detail[iType], -1);
          TLog.SelectDB('', Q);
          stLastCode := Q.FindField(stImmo).AsString;
          iMem := i;
          break;
        end;
      end;
    end
    else
    begin
      TLog := TOB.Create(stTable, T.Detail[iMem].Detail[iType], -1);
      TLog.SelectDB('', Q);
    end;
    Q.Next;
  end;
end;

procedure TPlanInfo.ChargeImmo(CodeImmo: string);
var
  i: integer;
  iTrouve: integer;
  Q: TQuery;
  stWhere, stSQL, stSelect: string;
  stOrderBy, stChamp, StValeur, stWhereV: string;
  T: TOB;
begin
  fDateCalcul := iDate1900;
  fTypeSortie := '';
  stOrderBy := '';
  stChamp := '';
  StValeur:= '';
  stWhereV := '';

  // Recherche de l'immobilisation dans le cache
  iTrouve := -1;
  for i := 0 to fTobImmo.Detail.Count - 1 do
  begin
    if fTobImmo.Detail[i].GetValue('I_IMMO') = CodeImmo then
    begin
      iTrouve := i;
      break;
    end;
  end;
  // L'immobilisation est chargée, on se positionne dessus.
  if iTrouve > -1 then
  begin
    fTobEnCours := fTobImmo.Detail[iTrouve];
  end
  else
    // L'immobilisation n'est pas chargée ==> on charge
  begin
    fTobImmo.ClearDetail;

    // BTY FQ 21169
    // 1- Revenir en arrière sur le tri de IMMO via SQLOrderBy i_compteimmo,i_immo
    // car tri IMMOLOG ou IMMOAMOR dans ChargeDetailImmo est i_immo
    // et on déroule IMMO-IMMOLOG ou IMMO-IMMOAMOR rang à rang pour lire les infos à calculer
    // => on sautait la lecture de IMMOLOG ou IMMOAMOR de certaines immos, d'où montants à zéro
    // 2- Conserver le SELECT ##TOP TAILLE_CACHE qui optimise la requête

    // BTY FQ 18370 2 Optimisations de la requête SQL :
    // 1. WHERE et ORDER BY = selon l'ordre de la boucle de traitement appelante
    // 2. SELECT ##TOP -> récupère uniquement le paquet TAILLE_CACHE
    // Exemple SQLWhere := '205000;0000012782;'
    //         SQLOrderBy := 'I_COMPTEIMMO;I_IMMO;'

    // stWhere := 'I_IMMO>="' + CodeImmo + '"';
    //Q := OpenSQL('SELECT * FROM IMMO WHERE ' + stWhere + ' ORDER BY I_IMMO',
    //     True);
   { if SQLOrderBy = '' then
    begin
      stOrderBy := 'I_IMMO';
      stWhere := 'I_IMMO>="' + CodeImmo + '"';
    end else
    begin
      // Démonter le Order By pour en extraire les champs
      // Démonter le Where pour en extraire les valeurs
      //stWhere := 'I_COMPTEIMMO+I_IMMO>="' + SQLWhere + CodeImmo + '"';
      while Length(SQLOrderBy) > 0 do
   	  begin
   	    stChamp := ReadTokenSt(SQLOrderBy);
        if stOrderBy = '' then
        begin
          stOrderBy := stChamp;
          stWhere := stChamp;
        end else
        begin
          stOrderBy := stOrderBy + ',' + stChamp;
          stWhere := stWhere + '+' + stChamp; // concaténer les champs
        end;
	      stValeur := ReadTokenSt(SQLWhere);
        //if (COPY (stChamp,1,6) = 'I_IMMO') or (COPY(stChamp,1,12) = 'I_COMPTEIMMO') then
        //   stValeur := stValeur +  StringOfChar('0', (17-Length(stValeur)+1));
        stWhereV := stWhereV + stValeur ;  // concaténer les valeurs des champs
      end;
      // Recomposer le Where
      stWhere := stWhere + '>="' + stWhereV + '"';
    end;

    Q := OpenSQL(
        'SELECT ##TOP '+IntToStr(TAILLE_CACHE)+'##  * FROM IMMO WHERE '+stWhere+' ORDER BY '+stOrderBy,
         True);
   }

    stWhere := 'I_IMMO>="' + CodeImmo + '"';
    Q := OpenSQL('SELECT ##TOP '+IntToStr(TAILLE_CACHE)+'##  * FROM IMMO WHERE ' + stWhere + ' ORDER BY I_IMMO',
         True);

    try
      i := 0;
      while not Q.Eof do
      begin
        Inc(i, 1);
        T := TOB.Create('IMMO', fTobImmo, -1);
        T.SelectDB('', Q, True);
        // Création de la mère du Log
        TOB.Create('LOG', T, -1);
        // Création de la mère du plan d'amortissement
        TOB.Create('AMORT', T, -1);
        if i = TAILLE_CACHE then
          break;
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;

    // Chargement du Log
    stWhere := ' WHERE ';
    for i := 0 to fTobImmo.Detail.Count - 1 do
    begin
      if Length(stWhere) > 7 then
        stWhere := stWhere + ' OR ';
      stWhere := stWhere + 'IL_IMMO="' + fTobImmo.Detail[i].GetValue('I_IMMO') +
        '"';
    end;
    // mbo fq 17512 ajout des zones double1 et double2 = cumantcesDR, CumantcesFEc
    // mbo fq 19712 ajout zone il_qtecedee
    // BTY 13/11/07 il_baseecoavmb pour récup base éco avant cession en édition IFR
    stSelect := 'SELECT IL_IMMO,IL_DATEOP,IL_TYPEOP,IL_TVAAREVERSER,IL_TVARECUPERABLE,IL_TVARECUPEREE,' +
      'IL_MONTANTDOT,IL_MONTANTCES,IL_VOCEDEE,IL_CUMANTCESECO,IL_CUMANTCESFIS, IL_DOUBLE1, IL_DOUBLE2, ' +
      'IL_PLANACTIFAP,IL_DUREEECO,IL_DUREEFISC,IL_METHODEECO,IL_METHODEFISC, IL_PVALUE, IL_CODEECLAT, IL_QTECEDEE, ' +
      'IL_BASEECOAVMB ';
    stSql := stSelect + 'FROM IMMOLOG ' + stWhere +
      ' ORDER BY IL_IMMO,IL_DATEOP ASC';
    Q := OpenSQL(stSQL, True);
    ChargeDetailImmo(fTobImmo, Q, INDICE_LOG);
    (*    while not Q.Eof do
        begin
          for i:=0 to fTobImmo.Detail.Count - 1 do
          begin
            if fTobImmo.Detail[i].GetValue('I_IMMO')=Q.FindField('IL_IMMO').AsString then
            begin
              T := TOB.Create ('IMMOLOG',fTobImmo.Detail[i].Detail[INDICE_LOG],-1);
              T.SelectDB('',Q);
              break;
            end;
          end;
          Q.Next;
        end;      *)
    Ferme(Q);
    // Chargement du plan d'amortissement
    stWhere := StringReplace(stWhere, 'IL_IMMO', 'IA_IMMO', [rfReplaceAll]);
    stSql := 'SELECT * FROM IMMOAMOR ' + stWhere +
      ' ORDER BY IA_IMMO,IA_NUMEROSEQ DESC,IA_DATE ASC';
    Q := OpenSQL(stSQL, True);
    (*    while not Q.Eof do
        begin
          for i:=0 to fTobImmo.Detail.Count - 1 do
          begin
            if fTobImmo.Detail[i].GetValue('I_IMMO')=Q.FindField('IA_IMMO').AsString then
            begin
              T := TOB.Create ('IMMOAMOR',fTobImmo.Detail[i].Detail[INDICE_AMORT],-1);
              T.SelectDB('',Q);
              break;
            end;
          end;
          Q.Next;
        end;      *)
    ChargeDetailImmo(fTobImmo, Q, INDICE_AMORT);
    Ferme(Q);
    if fTobImmo.Detail.Count > 0 then
      fTobEnCours := fTobImmo.Detail[0]
    else
      fTobEnCours := nil;
  end;
  if fTobEnCours <> nil then
  begin
    Plan.Init;
    fbSortie := (fTobEnCours.GetValue('I_OPECESSION') = 'X');
    bExcep := (fTobEnCours.GetValue('I_OPECHANGEPLAN') = 'X');
    fbEclate := (fTobEnCours.GetValue('I_OPEECLATEMENT') = 'X');
    fbPrime := (fTobEnCours.GetValue('I_SBVPRI') <> 0.00) or (fTobEnCours.GetValue('I_SBVPRIC') <> 0.00);
    SBV := (fTobEnCours.GetValue('I_SBVMT') <> 0.00) or (fTobEnCours.GetValue('I_SBVMTC') <> 0.00);

    Plan.ChargeTOB(fTobEnCours);
    // BTY 04/06 FQ 17629 Règle de calcul de la PMValue d'une cession
    RegleCession := fTobEnCours.GetValue('I_REGLECESSION');
  end;
end;

function TPlanInfo.GetValeurAchat(DateCalcul: TDateTime;
  bSortie: boolean): double;
begin
  if bSortie then
    Result := GetVoCedee(DateCalcul)+ Plan.ValeurTVARecuperable -
              Plan.ValeurTvaRecuperee  // fq 17292 ajout tva recuperable - tvarecuperee
  else if fbSortie then
    begin
    if fTobEnCours.GetValue('I_VALEURACHAT') = 0 then
      // cas des récup anciennes, I_VALEURACHAT=0, alors on se débrouille ...
      Result := fTobEnCours.GetValue('I_MONTANTBASEAMORT')
        // sinon on prend la valeur d'achat.
    else
      Result := fTobEnCours.GetValue('I_VALEURACHAT')+ Plan.ValeurTVARecuperable -
        Plan.ValeurTvaRecuperee;  // fq 17292 ajout tva recuperable - tvarecuperee
  end
  else
  begin
    { CA - 07/01/2004 - Dans le cas où l'immo a subi un éclatement, on récupère la
      valeur d'achat à la date de l'édition (date du plan)
      Dans les autres cas, on ne touche pas car ça marchait bien avant ...}
    { CA - 19/01/2004 - On utilise TableauDot dans le cas de l'éclatement pour
      retrouver la valeur de la reprise de l'immo avant éclatement }
    if fbEclate then
    begin
      if Plan.AmortEco.Reprise <> 0 then // Il y a une reprise sur l'immo
        //        Result := Plan.AmortEco.TableauDot[0]+Plan.AmortEco.BaseDebutExo[0]+Plan.ValeurTVARecuperable-Plan.ValeurTvaRecuperee
        Result := Plan.ValeurAchat + fValAchatEclate + Plan.ValeurTVARecuperable
          - Plan.ValeurTvaRecuperee
          //      else Result := Plan.AmortEco.BaseDebutExo[0]+Plan.ValeurTVARecuperable-Plan.ValeurTvaRecuperee;
      else
        Result := Plan.ValeurAchat + fValAchatEclate + Plan.ValeurTVARecuperable
          - Plan.ValeurTvaRecuperee;
    end
    else
      Result := Plan.ValeurHT + Plan.ValeurTVARecuperable -
        Plan.ValeurTvaRecuperee;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/09/2003
Modifié le ... :   /  /
Description .. : Calcul de la plus ou moins value à court terme
Mots clefs ... :
*****************************************************************}

function TPlanInfo.GetPVCT: double;
var
  CumulDotation: double;
begin
{  if CourtTerme then
    Result := PValue
  else
  begin
    if Plan.AmortEco.Methode = 'NAM' then
      Result := 0
    else
    begin
      if PValue < 0 then
        Result := PValue
      else
      begin
        CumulDotation := GetCumulAntEco(DateRef) + CessionEco;
        if PValue > CumulDotation then
          Result := CumulDotation
        else
          Result := PValue;
      end;
    end;
  end; }

  // BTY 04/06 FQ 17629 4 règles de calcul de la PMValue
  // CT  Tout en court terme
  // LT  Tout en long terme
  // NOR Reconduction existant
  // RSD Calcul normal en ignorant la durée (comme l'existant sans tenir compte de la durée)

  Result := PValue; // BTY 05/06 Pour éviter messages compilateur
  // Tout PMValue se trouve en court terme
  if RegleCession = 'CT' then
     begin
     Result := PValue;
     end
  // Tout PMValue se trouve en long terme
  else if RegleCession = 'LT' then
     begin
     Result := 0;
     end
  // Soit calcul normal, soit le normal en ignorant la durée
  else if (RegleCession = '') or (RegleCession = 'NOR') or (RegleCession = 'RSD') then
     begin
     if CourtTerme and (not (RegleCession = 'RSD')) then
        Result := PValue
     else
        begin
        if Plan.AmortEco.Methode = 'NAM' then
           Result := 0
        else
           begin
           if PValue < 0 then
              Result := PValue
           else
              begin
              CumulDotation := GetCumulAntEco(DateRef) + CessionEco;
              if PValue > CumulDotation then
                 Result := CumulDotation
              else
                 Result := PValue;
              end;
           end; // Plan.AmortEco.Methode
        end;
     end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Calcul de la plus ou moins value à long terme
Mots clefs ... :
*****************************************************************}

function TPlanInfo.GetPVLT: double;
var
  CumulDotation: double;
begin
{  if CourtTerme then
    Result := 0
  else
  begin
    Result := 0;
    if Plan.AmortEco.Methode = 'NAM' then
      Result := PValue
    else
    begin
      CumulDotation := GetCumulAntEco(DateRef) + CessionEco;
      if PValue < 0 then
        Result := 0
      else if (PValue > CumulDotation) then
        Result := PValue - CumulDotation;
    end;
  end; }
  // BTY 04/06 FQ 17629 4 règles de calcul de la PMValue
  // CT  Tout en court terme
  // LT  Tout en long terme
  // NOR Reconduction existant
  // RSD Calcul normal en ignorant la durée (comme l'existant sans tenir compte de la durée)

  Result := 0;     // BTY 05/06 Pour éviter messages compilateur
  // Tout PMValue se trouve en court terme
  if RegleCession = 'CT' then
     begin
     Result := 0;
     end
  // Tout PMValue se trouve en long terme
  else if RegleCession = 'LT' then
     begin
     Result := PValue;
     end
  // Soit calcul normal, soit le normal en ignorant la durée
  else if (RegleCession = '') or (RegleCession = 'NOR') or (RegleCession = 'RSD') then
     begin
     if CourtTerme and (not (RegleCession = 'RSD')) then
        Result := 0
     else
        begin
        Result := 0;
        if Plan.AmortEco.Methode = 'NAM' then
           Result := PValue
        else
           begin
           CumulDotation := GetCumulAntEco(DateRef) + CessionEco;
           if PValue < 0 then
              Result := 0
           else if (PValue > CumulDotation) then
              Result := PValue - CumulDotation;
           end;
        end;
     end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/09/2003
Modifié le ... :   /  /
Description .. : Détermine si la plus value est à court terme ou long terme
Mots clefs ... :
*****************************************************************}

function TPlanInfo.CourtTerme: boolean;
begin
  Result := fDateSortie < PlusMois(Plan.DateAchat, 24);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 21/01/2004
Modifié le ... :   /  /
Description .. : Renvoie la reprise totale d'une immobilisation qui a subi un
Suite ........ : éclatement avant l'éclatement en allant chercher la reprise
Suite ........ : de la partie éclatée
Mots clefs ... :
*****************************************************************}

procedure TPlanInfo.GetRepriseSuiteEclatement(
  DateCalcul: TDateTime; var RepriseEco: double; var RepriseFiscale: double);
var
  i: integer;
  T: TOB;
  stCodeEclate: string;
  Q: TQuery;
begin
  fValAchatEclate := 0;
  stCodeEclate := '';
  RepriseEco := 0;
  RepriseFiscale := 0;
  if fTobEnCours <> nil then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      if ((T.GetValue('IL_DATEOP') > DateCalcul) and (T.GetValue('IL_TYPEOP') =
        'ECL')) then
        stCodeEclate := T.GetValue('IL_CODEECLAT');
    end;
    if stCodeEclate <> '' then
      // un éclatement a eu lieu après la date de calcul (la reprise est erronée )
    begin
      Q :=
        OpenSQL('SELECT I_REPRISEECO, I_REPRISEFISCAL, I_VALEURACHAT, I_REPCEDECO, I_REPCEDFISC FROM IMMO WHERE I_IMMO="'
        + stCodeEclate + '"', True);
      if not Q.Eof then
      begin
        RepriseEco := Q.FindField('I_REPRISEECO').AsFloat +
          Q.FindField('I_REPCEDECO').AsFloat;
        RepriseFiscale := Q.FindField('I_REPRISEFISCAL').AsFloat +
          Q.FindField('I_REPCEDFISC').AsFloat;
        fValAchatEclate := Q.FindField('I_VALEURACHAT').AsFloat;
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Reprise des antérieurs dérogatoires et réintégration
Mots clefs ... :
*****************************************************************}
procedure TPlanInfo.GetRepriseDREclatement(
  DateCalcul: TDateTime; var RepriseDR: double; var RepriseFEC: double);
var
  i: integer;
  T: TOB;
  stCodeEclate: string;
  Q: TQuery;
begin
  fValAchatEclate := 0;
  stCodeEclate := '';
  RepriseDR := 0;
  RepriseFEC := 0;
  if fTobEnCours <> nil then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      if ((T.GetValue('IL_DATEOP') > DateCalcul) and (T.GetValue('IL_TYPEOP') =
        'ECL')) then
        stCodeEclate := T.GetValue('IL_CODEECLAT');
    end;
    if stCodeEclate <> '' then
      // un éclatement a eu lieu après la date de calcul (la reprise est erronée )
    begin
      Q :=
        OpenSQL('SELECT I_REPRISEDR, I_REPRISEFEC, I_REPRISEFDRCEDEE, I_REPRISEFECCEDEE FROM IMMO WHERE I_IMMO="'
        + stCodeEclate + '"', True);
      if not Q.Eof then
      begin
        RepriseDR := Q.FindField('I_REPRISEDR').AsFloat +
          Q.FindField('I_REPRISEFDRCEDEE').AsFloat;
        RepriseFEC := Q.FindField('I_REPRISEFEC').AsFloat +
          Q.FindField('I_REPRISEFECCEDEE').AsFloat;
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 05/09/2006
Modifié le ... :   /  /
Description .. : Récupère les antérieurs Prime suite à un éclatement
Mots clefs ... :
*****************************************************************}
procedure TPlanInfo.GetReprisePriEclatement(DateCalcul: TDateTime; var ReprisePri: double);
var
  i: integer;
  T: TOB;
  stCodeEclate: string;
  Q: TQuery;
begin
  fValAchatEclate := 0;
  stCodeEclate := '';
  ReprisePri := 0;
  if fTobEnCours <> nil then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      if ((T.GetValue('IL_DATEOP') > DateCalcul) and (T.GetValue('IL_TYPEOP') =
        'ECL')) then
        stCodeEclate := T.GetValue('IL_CODEECLAT');
    end;
    if stCodeEclate <> '' then
      // un éclatement a eu lieu après la date de calcul (la reprise est erronée )
    begin
      Q :=
        OpenSQL('SELECT I_REPRISEUO, I_VALEURACHAT, I_REPRISEUOCEDEE FROM IMMO WHERE I_IMMO="'
        + stCodeEclate + '"', True);
      if not Q.Eof then
      begin
        ReprisePri := Q.FindField('I_REPRISEUO').AsFloat +
          Q.FindField('I_REPRISEUOCEDEE').AsFloat;
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 29/09/2006
Modifié le ... : 19/10/2006
Description .. : Récupère les antérieurs SUBVENTION suite à un
Suite ........ : éclatement si l'éclatement a eu lieu après la date de calcul
Mots clefs ... :
*****************************************************************}
procedure TPlanInfo.GetRepriseSbvEclatement(DateCalcul: TDateTime; var RepriseSbv: double);
var
  i: integer;
  T: TOB;
  stCodeEclate: string;
  Q: TQuery;
begin
  fValAchatEclate := 0;
  stCodeEclate := '';
  RepriseSbv := 0;
  if fTobEnCours <> nil then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
      if ((T.GetValue('IL_DATEOP') > DateCalcul) and (T.GetValue('IL_TYPEOP') =
        'ECL')) then
        stCodeEclate := T.GetValue('IL_CODEECLAT');
    end;
    if stCodeEclate <> '' then
      // un éclatement a eu lieu après la date de calcul (la reprise est erronée )
    begin
      Q :=
        OpenSQL('SELECT I_CORRECTIONVR FROM IMMO WHERE I_IMMO="'
        + stCodeEclate + '"', True);
      if not Q.Eof then
      begin
        RepriseSbv := Q.FindField('I_CORRECTIONVR').AsFloat;
      end;
    end;
  end;
end;


function TPlanInfo.DateDebutDotation(DateRef: TDateTime): TDateTime;
var
  Exo: TExoDate;
begin
{$IFDEF SERIE1}
  if ImQuelDateDeExo(ImQuelExoDt(DateRef), Exo) then
    Result := Exo.Deb
  else
    Result := 0;
{$ELSE}
{$IFDEF EAGLSERVER}
  if CQuelExerciceServer(Exercices, DateRef, Exo)
{$ELSE}
  if CQuelExercice(DateRef, Exo)
{$ENDIF} then
    Result := Exo.Deb
  else
    Result := 0;
{$ENDIF}
  if Result = 0 then
    Result := EvalueDateExercice(DateRef, False);
end;

function TPlanInfo.DateFinDotation(DateRef: TDateTime): TDateTime;
var
  Exo: TExoDate;
begin
{$IFDEF SERIE1}
  if ImQuelDateDeExo(ImQuelExoDt(DateRef), Exo) then
    Result := Exo.Fin
  else
    Result := 0;
{$ELSE}
{$IFDEF EAGLSERVER}
  if CQuelExerciceServer(Exercices, DateRef, Exo)
{$ELSE}
  if CQuelExercice(DateRef, Exo)
{$ENDIF} then
    Result := Exo.Fin
  else
    Result := 0;
{$ENDIF}
  if Result = 0 then
    Result := EvalueDateExercice(DateRef, True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/08/2004
Modifié le ... :   /  /
Description .. : Cette fonction évalue une date de fin ou de début
Suite ........ : d'exercice dans le cas ou aucun exercice ne correspond
Suite ........ : dans la base à la date passée en paramètre
Mots clefs ... :
*****************************************************************}

function TPlanInfo.EvalueDateExercice(DateRef: TDateTime; bFin: boolean):
  TDateTime;
var
  DateTmp: tDateTime;
{$IFDEF EAGLSERVER}
  ExoSuivant: TExoDate;
{$ENDIF}
begin
  if bFin then
  begin
{$IFDEF EAGLSERVER}
    ExoSuivant := SuivantServer;
    if ((ExoSuivant.Code <> '') and (DateRef<=ExoSuivant.Fin))then
      DateTmp := ExoSuivant.Fin
    else
      DateTmp := FinDeMois(PlusMois(EncoursServer.Fin, 12));
{$ELSE}
    if ((VHImmo^.Suivant.Code <> '') and (DateRef<=VHImmo^.Suivant.Fin)) then
      DateTmp := VHImmo^.Suivant.Fin
    else DateTmp := FinDeMois(PlusMois(VHImmo^.Encours.Fin, 12));
{$ENDIF}
     while ( DateTmp < DateRef ) do
      DateTmp := PlusMois(DateTmp, 12);
  end
  else
  begin
{$IFDEF EAGLSERVER}
    ExoSuivant := SuivantServer;
    if ExoSuivant.Code <> '' then
      DateTmp := ExoSuivant.Deb
    else
      DateTmp := EncoursServer.Fin + 1;
{$ELSE}
    if VHImmo^.Suivant.Code <> '' then
      DateTmp := VHImmo^.Suivant.Deb
    else
      DateTmp := VHImmo^.Encours.Fin + 1;
{$ENDIF}
    while (PlusMois(DateTmp, 12) <= DateRef) do
      DateTmp := PlusMois(DateTmp, 12);
  end;
  Result := DateTmp;
end;

{$IFDEF EAGLSERVER}

function CQuelExerciceServer(Exercices: array of TExoDate; Date: TDateTime; var
  Exo: TExoDate): boolean;
var
  i: integer;
begin
  i := 0;  //1;  BTY 02/06 FQ 17547
  Result := False;
  while (Exercices[i].Code <> '') do
  begin
    if (Date >= Exercices[i].Deb) and (Date <= Exercices[i].Fin) then
    begin
      Exo := Exercices[i];
      Result := true;
      break;
    end;
    Inc(i, 1);
  end;
end;

procedure TPlanInfo.SetExercicesServer(Exo: array of TExoDate);
var
  i: integer;
begin
  //BTY 02/06 FQ 17547
  //for i := 1 to 20 do
  for i := 0 to 20 do
    Exercices[i] := Exo[i];
  Plan.SetDatesServer(EncoursServer);
end;

procedure TPlanInfo.UpdateInfo(CodeImmo: string; LaDate: TDateTime);
begin
  if (Plan.CodeImmo <> CodeImmo) or (LaDate <> DateRef) then
  begin
    ChargeImmo(CodeImmo);
    Calcul(LaDate, true, false);
  end;
end;

function TPlanInfo.EncoursServer: TExoDate;
var
  i: integer;
begin
  //  FillChar(Result,SizeOF(Exercices),#0) ;
  //BTY 02/06 FQ 17547
  //BTY 03/06 Attention exercice en cours peut être CPR cf VHIMMO
  //for i := 1 to 20 do
  for i := 0 to 20 do
  begin
   // if (Exercices[i].EtatCpta = 'OUV') then
    if (Exercices[i].EtatCpta = 'OUV') or
       (Exercices[i].EtatCpta = 'CPR') then
    begin
      Result := Exercices[i];
      break;
    end;
  end;
end;

function TPlanInfo.SuivantServer: TExoDate;
var
  i: integer;
begin
  //  FillChar(Result,SizeOF(Exercices),#0) ;
  //BTY 02/06 FQ 17547
  //for i := 1 to 20 do
  for i := 0 to 20 do
  begin
    if (Exercices[i].EtatCpta = 'OUV') then
    begin
      Result := Exercices[i + 1];
      break;
    end;
  end;
end;

// BTY 02/06 Paramètres TVA accessibles aux éditions via le contexte PlanInfo
procedure TPlanInfo.SetTauxTVAServer(TobTvaWA : TOB);
begin
  TobTva := TOB.Create('',Nil,-1);
  // Charger la tob de TPlanInfo avec celle passée en entrée
  if TobTvaWA <> nil then
     TobTva.Dupliquer(TobTvaWA,True,True);
end;

// BTY 03/06 Coefficients dégressifs accessibles aux éditions via le contexte PlanInfo
procedure TPlanInfo.SetCoeffDegressifServer(CoeffDegressifWA : TOB);
begin
  CoeffDegressif := TOB.Create('',Nil,-1);
  // Charger la tob de TPlanInfo avec celle passée en entrée
  if CoeffDegressifWA <> nil then
     CoeffDegressif.Dupliquer(CoeffDegressifWA,True,True);
end;

{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 04/01/2005
Modifié le ... :   /  /
Description .. : Calcul du dérogatoire
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetDerogatoire: double;
var
  CE, CF, AE, AF : double;
  dDerog, dReint : double;
  Tva, PV : double;
  CumulDerog, CumulReint :double;
begin
//  if GetInfoCession(DateRef, Tva, PV) then Result := (-1)*GetCumulAntDerogatoire
//  else
  begin
    CumulDerog := GetCumulAntDerogatoire(DateRef);
    CumulReint := GetCumulAntFEC(DateRef);

    AE := GetCumulAntEco(DateRef);
    AF := GetCumulAntFisc(DateRef);


    CE := AE+DotationEco;

    if GetInfoCession(DateRef, Tva, PV) then CF := AF + CessionFisc
    else CF := AF+DotationFisc;

    dDerog := 0;
    dReint := 0;


    CalculDerogatoire(DotationEco, DotationFisc, AE, AF, CE, CF, CumulDerog, CumulReint, dDerog, dReint);

    Result := dDerog;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 09/03/2007
Modifié le ... :   /  /    
Description .. : Calcul de la réintégration fiscale (réintégration ou déduction)
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetReintegration: double;
var
  CE, CF, AE, AF : double;
  dDerog, dReint : double;
  Tva, PV : double;
  CumulDerog, CumulReint :double;
begin
//  if GetInfoCession(DateRef, Tva, PV) then Result := (-1)*GetCumulAntDerogatoire
//  else
  begin
    CumulDerog := GetCumulAntDerogatoire(DateRef);
    CumulReint := GetCumulAntFEC(DateRef);

    AE := GetCumulAntEco(DateRef);
    AF := GetCumulAntFisc(DateRef);


    CE := AE+DotationEco;

    if GetInfoCession(DateRef, Tva, PV) then CF := AF + CessionFisc
    else CF := AF+DotationFisc;

    CalculDerogatoire(DotationEco, DotationFisc, AE, AF, CE, CF, CumulDerog, CumulReint, dDerog, dReint);

    Result := dReint;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : renvoie le cumul ANTERIEUR dérogatoire pratiqué
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetCumulAntDerogatoire(DateRef : TDateTime): double;
var
  i : integer;
begin
  i := 0;
  Result := Plan.AmortDerog.Reprise + Plan.AmortDerog.RepriseCedee;

  while ((Plan.TableauDate[i]<>iDate1900)) and (Plan.TableauDate[i]<DateDebutDotation(DateRef)) do
  begin
    Result := Result  + Plan.AmortDerog.TableauDot[i] +Plan.AmortDerog.TabCession[i];
    Inc (i);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : renvoie le cumul dérogatoire pratiqué lors de la cession
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetCumulDerogatoireCede : double;
var
  i : integer;
begin
  i := 0;
  Result := Plan.AmortDerog.Reprise + Plan.AmortDerog.RepriseCedee;

  while ((Plan.TableauDate[i]<>iDate1900)) do
  begin
    Result := Result  + Plan.AmortDerog.TableauDot[i];
    Inc (i);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : renvoi le cumul anterieurs des réintégrations/Déductions à une date
Suite ........ : donnée
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetCumulAntFEC(DateRef : TDateTime) : double;
var
  i : integer;
begin
  i := 0;
  Result := Plan.AmortReint.Reprise + Plan.AmortReint.RepriseCedee;

  while ((Plan.TableauDate[i]<>iDate1900)) and (Plan.TableauDate[i]<DateDebutDotation(DateRef)) do
  begin
    Result := Result  + Plan.AmortReint.TableauDot[i] +Plan.AmortReint.TabCession[i];
    Inc (i);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : renvoi le cumul des réintégrations/Déductions cedé
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetCumulFECcede : double;
var
  i : integer;
begin
  i := 0;
  Result := Plan.AmortReint.Reprise + Plan.AmortReint.RepriseCedee;

  while ((Plan.TableauDate[i]<>iDate1900)) do
  begin
    Result := Result  + Plan.AmortReint.TableauDot[i];
    Inc (i);
  end;
end;


//=============================================================

procedure TPlanInfo.CalculDerogatoire(DotEco, DotFisc, AE, AF, CE, CF, CumulDerog, CumulReint: double;
                                      var DotDerog, DotReint : double);
var EC, EA : double;
begin
  { FQ 15729 & 15728 - Ajout Arrondi }
 { EC := Arrondi (CF-CE, V_PGI.OkDecV);
  EA := Arrondi (AF-AE, V_PGI.OkDecV);
  Dotation := 0;
  Reprise := 0;
  if (EC >= 0) then
  begin
    if EA < 0 then EA := 0;
    if EC > EA then Dotation := EC-EA
    else if EA > EC then Reprise := EA-EC;
  end else
  if (EC < 0) then if (EA>0) then Reprise := EA;}

  EC := arrondi((CF-CE),V_PGI.OkDecV); //écart sur (cumul antérieur + dotation de l'exercice)
  EA := arrondi((AF-AE),V_PGI.OkDecV); //écart sur cumuls antérieurs

  DotDerog := 0;
  DotReint := 0;

  if Plan.AmortEco.Methode = 'VAR' then
  begin

     if not (Plan.GestionFiscale) then
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
        // mais pas l'inverse : c'est à dire qu'on se préoccupe pas du cumul réintégration avant de
        // générer du dérogatoire
        if (DotFisc > DotEco) then
        begin
           DotDerog := arrondi((DotFisc - DotEco),V_PGI.OkDecV);
           DotReint := 0;
        end else
        begin
           if CumulDerog = 0 then
           begin
              DotReint := (DotFisc - DotEco) * -1;
              DotReint := arrondi(DotReint,V_PGI.OkDecV);
           end else
           begin
              if CumulDerog > arrondi(((DotFisc - DotEco)* -1),V_PGI.OkDecV) then
              begin
                DotDerog := arrondi((DotFisc - DotEco),V_PGI.OkDecV);  // on est en reprise de dérogatoire
              end else
              begin
                // on va solder le dérog et générer de la réintégration
                DotDerog := CumulDerog * -1;
                DotReint := (CumulDerog + DotFisc - DotEco)*-1;
                DotDerog := arrondi(DotDerog,V_PGI.OkDecV);
                DotReint := arrondi(DotReint,V_PGI.OkDecV);
              end;
           end;
        end;
     end;
  end else
  begin
       // le plan d'amortissement éco n'est pas variable
     if not (Plan.GestionFiscale) then
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
          DotDerog := 0;
     end;
  end;

end;
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/01/2005
Modifié le ... :   /  /
Description .. : Dotation réellement pratiquée = dotation économique +
Suite ........ : dérogatoire
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetDotationPratiquee: double;
begin
  Result := DotationEco + Derogatoire;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/01/2005
Modifié le ... :   /  /
Description .. : Calcul du cumul antérieur réellement pratiqué = Cumul
Suite ........ : antérieur économique + cumul antérieur dérogatoire
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetCumulAntPratique: double;
begin
  Result := CumulAntEco + GetCumulAntDerogatoire(DateRef);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/01/2005
Modifié le ... :   /  /
Description .. : Calcul de la dotation fiscale réellement pratiquée lors d'une
Suite ........ : sortie
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetCessionFiscPratiquee: double;
begin
  Result := CessionEco + Derogatoire;
end;
{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 18/11/2005
Modifié le ... :   /  /
Description .. : récupere le montant de la dépréciation en fonction d'une
Suite ........ : date donnée (utilisée dans les éditions)
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetDotDeprec(DateCalcul: TDateTime): double;
var
  i: integer;
  T: TOB;
  // suppression mbo 10.01.2006 car non utilisée Q: TQuery;
begin
  result := 0;
  if fTobEnCours <> nil then
  begin
    for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
    begin
      T := fTobEnCours.Detail[INDICE_LOG].Detail[i];

// BTY 03/06 Pb en édition Web Access avec VHIMMO
{$IFDEF EAGLSERVER}
      if ((T.GetValue('IL_DATEOP') <= DateCalcul) and (T.GetValue('IL_TYPEOP') =
        'DPR') and (T.GetValue('IL_DATEOP') >= EnCoursServer.Deb)
        and (T.GetValue('IL_DATEOP') <= EnCoursServer.Fin))then
{$ELSE}
      if ((T.GetValue('IL_DATEOP') <= DateCalcul) and (T.GetValue('IL_TYPEOP') =
        'DPR') and (T.GetValue('IL_DATEOP') >= VHImmo^.EnCours.Deb)
        and (T.GetValue('IL_DATEOP') <= VHImmo^.EnCours.Fin))then
{$ENDIF}
        result := T.GetValue('IL_MONTANTDOT');
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 13/11/2007
Modifié le ... :   /  /
Description .. : Fonction pour l'édition IFR des immos issues du régime forfait
Suite ........ : Récup VNC restant à amortir pour un bien agricole
Suite ........ : VNC restant à amortir = base éco de l'immo une fois le passage au réel validé
Suite ........ : Si immo cédée, récup de sa base éco avant cession
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetVNCIFR(DateCalcul: TDateTime; bSortie: boolean): double;
var i: integer;
    T : Tob;
begin
  if Plan.AmortEco.Methode = 'NAM' then
    result := 0
  else
  begin
    // si sortie avant la date arrêté (bSortie) ou après (fbSortie)
    if bSortie or fbSortie then
    begin
      result := 0;
      for i := 0 to fTobEnCours.Detail[INDICE_LOG].Detail.Count - 1 do
      begin
        T := fTobEnCours.Detail[INDICE_LOG].Detail[i];
        if T.GetValue('IL_TYPEOP') = 'CES' then
        begin
          result := T.GetValue('IL_BASEECOAVMB');
          Break;
        end;
      end;
    end else
    begin
      result := Plan.AmortEco.Base;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 14/11/2007
Modifié le ... :   /  /
Description .. : Fonction pour l'édition IFR des immos issues du régime forfait
Suite ........ : Récup durée restant à amortir d'un bien agircole
Suite ........ : Durée = durée éco de l'immo une fois le passage au réel validé
Suite ........ : Même chose si immo cédée
Mots clefs ... :
*****************************************************************}
function TPlanInfo.GetDureeIFR(DateCalcul: TDateTime): double;
begin
  if Plan.AmortEco.Methode = 'NAM' then
    result := 0
  else
    result := Plan.AmortEco.Duree;
end;
//------------------------------------------------------------------------------

end.




