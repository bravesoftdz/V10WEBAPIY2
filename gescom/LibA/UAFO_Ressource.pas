{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... : 14/09/2001
Description .. : OBJET RESSOURCE
Mots clefs ... : OBJET; RESSOURCE
*****************************************************************}
unit UAFO_Ressource;

interface

uses
  {$IFDEF VER150} Variants, {$ENDIF}
  Classes,
{$IFDEF EAGLCLIENT}
{$ELSE}
  db,
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}
{$ENDIF}
  sysutils,
  paramsoc,

{$IFDEF AFFAIRE}
  afplanningCst,
  ed_tools,
  CalcOleGenericAff,
  stdctrls,
  hPanel,
  UAFO_REGLES,
  dicoBTP,
{$ENDIF}
{$IFDEF PLUGIN}
  uToolsPlugin,
{$ENDIF !PLUGIN}
  UAFO_FERIE,
  HCtrls,
  HEnt1,
  utob,
  formule
  ;

(******************************************************************************)
// TAFO_Ressource : objet image d'une ressource pouvant être créé à partir
// d'un code ressource. Seul le code est obligatoire pour la création d'une
// entité Ressource.
//
// services :
//
// FonctionDeLaRessource : si une date est passée en entrée, renvoie une TStringList
//                      comprenant un seul élément : la fonction en cours pour la date donnée
//                         si aucune date n'est fournie, renvoie une TStringList,
//                      comprenant les 4 dernières fonctions occupées triées de la plus
//                      récente (indice 0) à la plus ancienne (indice 3).
// PrixCourant : si une date est passée en entrée, renvoie une TStringList
//                      comprenant un seul élément : le prix en cours pour la date donnée
//                         si aucune date n'est fournie, renvoie une TStringList,
//                      comprenant les 4 derniers prix triés du plus
//                      récent (indice 0) au plus ancien (indice 3).
// PRRessource : renvoie le prix de revient de la ressource
// PVRessource : renvoie le prix de vente HT ou TTC de la ressource
// RessourceDisponible : renvoie un booleen indiquant si la ressource est disponible
//                      à la date et heure entrée
// ListeCompetencesD1Ressource : renvoie une TStringList des compétences de la ressource
// AssocieArticle : associe un article à la ressource. Si pas d'article en entree, associe
//                   l'article par defaut de la ressource
// Auteur : Pierre Lenormand le 02/11/99
// PlugIn : Vignette portail : Ressource en sous-charge
(******************************************************************************)
type
  TAFO_Ressource = class
    Ftob_Champs: TOB;

  private
{$IFDEF AFFAIRE}
    //fMMRapportDec : TMemo;               
    //fMMRapportAnn : TMemo;
    fSlRapportDec: TStringList;
    fSlRapportAnn: TStringList;
    procedure ControleListeEtDecale(var pListeJours: TListDate; pRdDuree: Double; pInNbJoursDecal: Integer; pStMethodeDecal: string; pTobJoursOccupes: Tob; pStAffaire, pStRes: string; pTache: RecordTache);
    procedure ControleLastDateGene(var pListeJours: TListDate; pLastDateGene: TdateTime; pBoAjout: Boolean);
    function DecalageJour(var pListeJours: TListDate; pInNbJoursDecal: Integer; pStMethodeDecal: string; pRdDuree: Double; var pStJourEnCours: TdateTime; pTobJoursOccupes: Tob): Integer;
    function PeriodeTravaillee(pStCurJour: TdateTime; pRdDuree: Double; var pStJourTrouve: TdateTime): Boolean;
    function LoadListeJoursOccupes(pStRes: string; pDtDateDeb, pDtDateFin: TDateTime): Tob;
    function PasDansListe(vDtCurDate: TDateTime; pSLJours: TListDate; pBoAjout: Boolean): Boolean;
    function JourPasOccupe(vDtCurDate: TDateTime; pTobJoursOccupes: Tob): Boolean;
    //procedure AddJournal(pStType, PStAffaire, pStTiers, pStInterv, pStRes, pStDtSource, pStDtDest : String);
    //procedure CreeJournaux;

    function GenererLignesPlanning(pListeJours: TListDate; var pInNumLigne: Integer; pTob: Tob; pStAffaire: string;
      pTache: RecordTache; pArticle: RecordArticle; pBoValorisation: Boolean; pRdDuree: Double): Boolean;

{$ENDIF}
    function DonneesFormuleP(AcsChampFormule: Hstring): variant;
    function PaieUtilisee: Boolean;

  public
    Article: TOB; // Variable servant à faire pointer l'occurrence de la ressource sur un article
    TobRessourcePR: TOB;
    TobCalendrier: TOB;
    TobCalendrierRegle: TOB;
    TobAbsenceSalarie: TOB; // Absences saisies dans la paie
    TobAbsencesParPeriode: TOB; // Tob de toutes les absences déjà calculées rangées par période pour la ressource
    TobDisposParPeriode: TOB; // Tob de toutes les dispos déjà calculées rangées par période pour la ressource
    gsPeriodiciteCourante: string; // Periodicite qui a servi à calculer les absences (sert à savoir si on doit recalculer au cas où la périodicité a changé)
    gsUniteValo: string;
    gbPVHT: boolean;
    gDatePrix: TDateTime;
    gdTempsAnnuel: double;
    giAbsencePaie: integer; // 0=jamais initialisé; 1=les absences sont gérées dans la paie pour la ressource courante; -1=pas d'absence pour ce salarié dans la paie
    gbCreateLeger: boolean; // create léger pour prévenir que seuls quelques champs ont été chargés à la création
    gbCreateAvecLibres: boolean; // create légère avec les champs libres en +
    gobj_JourFerie: TJourFerie;
    gsUniteActivite :string;  //AB-200607-
{$IFDEF PLUGIN}
    plugin :TVignettePlugin;
    constructor Create(AcsCodeRessource: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false;APlugIn :TVignettePlugIn=nil); overload;
{$ELSE}
    constructor Create(AcsCodeRessource: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false); overload;
{$ENDIF}
    destructor Destroy; override;
    property tob_Champs: TOB read Ftob_Champs write Ftob_Champs;

    function FonctionDeLaRessource(AcdDateDeFonction: TDateTime): TStringList;
    function PrixCourant(AcdDateDePrix: TDateTime; AcsTypePrix: string; AcbPVHT: boolean = true): TStringList;
    function PRRessource(AcdDate: TDateTime; ActobArticle: TOB; ActobAffaires: TOB; AcsCodeArticle: string; AcsCodeAffaire: string = ''): double;
    function PVRessource(AcdDate: TDateTime; AcbHT: boolean; ActobArticle: TOB; ActobAffaires: TOB; AcsCodeArticle: string; AcsCodeAffaire: string = ''): double;
    function CalculPrixRessource(AcdDate: TDateTime; AcsTypePrix: string; ActobArticle: TOB; ActobAffaires: TOB; AcsCodeArticle: string; AcsCodeAffaire: string = ''): double;
    function RessourceDisponible(AcdDateHeureDemande: TDateTime): boolean;
    function ListeCompetencesD1Ressource: TStringList;
    function AssocieRessourcePR(AcsCodeRessource: string; AcbForcer: boolean): boolean;
    procedure ChargeCalendriers(AcbForcer: boolean);
//    function AUnCalendrier : boolean;
    procedure ChargeAbsencesSalarie(AcbForcer: boolean);
    function ChargeTousLesChamps(AcsRessource: string): boolean;
    procedure ChargeTempsAnnuel35H;
    function TempsAnnuelRessource(AcsTypeActivite: string; AcdDateDebut, AcdDateFin: TDateTime): double;
    function ExisteUnArticleDeValo(AcsTypeValo, AcsArticle: string): boolean;
    function DetDerniereDateEffetCalendrier(AcdDateRef: TDateTime; var AviIndice: integer): TDateTime;
    function AbsenceSalarieD1Date(AcdDateTestee: TDateTime): TOB;
    function JourTravaille(pDtWork: TDateTime; var NbHeure: double; bAxeHeures: boolean): Boolean;
    function NbHeuresDispo(pDtDebut, pDtFin: TDateTime; sPeriode: string; bForcer: boolean = false): TOB;
    procedure ChargeTOBDispoDeLaRessource(TOBMere: TOB; pDtDebut, pDtFin: TDateTime; pPeriodicite: string);
{$IFDEF AFFAIRE}
    function AssocieArticle(AcsCodeArticle: string): boolean;
    function GenererPlanning(var pInNumLigne: Integer; pTob: Tob; pAFReglesTache: TAFRegles;
      pStAffaire: string; pTache: RecordTache; pArticle: RecordArticle;
      pBoValorisation, pBoAjout: Boolean; pPanel: ThPanel): Boolean;
    //function GetRessourceRemplacement : TAFO_Ressource;
    function NbJoursAbsence(pDtDebut, pDtFin: TDateTime; sPeriode: string; bForcer: boolean = false): TOB;
    procedure ChargeTOBAbsenceDeLaRessource(TOBMere: TOB; pDtDebut, pDtFin: TDateTime; pPeriodicite: string);
{$ENDIF}
  end;

  (******************************************************************************)
  // TAFO_Ressources : liste d'elements de type TAFO_Ressource hérité du type TSTRINGLIST
  // Sert à la gestion par listes des ressources
  //
  // services :
  // RessourcesD1Competence : donne la liste des ressources possedant une competence
  // RessourcesD1Fonction : donne la liste des ressources ayant eu ou ayant actuellement
  //                         une fonction. Pour savoir si une des fonctions est valable
  //                         à une date donnée, il faut appeler le service 'FonctionDeLaRessource'
  //                         ressource par ressource.
  //
  // Libération : la mémoire est entièrement libérée, par ressource et pour la liste
  //                elle même.
  //
  // Utilisation courante :
  //       on créé une occurrence de TAFO_Ressources, un appel à l'un des services
  //       rempli la liste, la propriété count indiquant le nombre d'entité TAFO_Ressources
  //       existant dans la liste. On peut aussi gérer complètement sa liste de TAFO_Ressources
  //       en utilisant toutes les méthodes de l'ancêtre TLIST.
  //
  // Auteur : Pierre Lenormand le 02/11/99
  (******************************************************************************)
type
  TAFO_Ressources = class(TStringList)

  public
    destructor Destroy; override;

    function AddRessource(AcsRessource: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false): integer;
    function TOBRessources(AcsTitreTOB: string): TOB;
    function GetRessource(Index: Integer): TAFO_Ressource;

    procedure RessourcesD1Competence(AcsCompetence: string);
    procedure RessourcesD1Fonction(AcsFonction: string);
    function RessourceD1User(AcsUser: string): string;
		function AddRessourceD1User(AcsUser: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false): integer;
{$IFDEF AFFAIRE}
    procedure RessourceD1TacheAffaire(AcsAffaire, AcsNumeroTache: string);
    function TOBAbsencesDesRessources(pDtDebut, pDtFin: TDateTime; pPeriodicite: string; TobRessource: TOB = nil): TOB;
    function TOBDisponibiliteDesRessources(pDtDebut, pDtFin: TDateTime; pPeriodicite: string; TobRessource: TOB = nil): TOB;
{$ENDIF}
  end;

function RessourceDuSalarie(AcsCodeSalarie: string): string;

implementation

uses
{$IFDEF AFFAIRE}
  UtilTaches,
  AffaireUtil,
  AFUtilArticle,
  utilGa,
{$ENDIF}
  entgc,
  UtilRessource
  ;

const
  TexteMessage: array[1..8] of string = (
    {1}'Affaire : %s. Tiers : %s. Intervention %s non générée le %s pour la ressource %s.',
    {2}'Affaire : %s. Tiers : %s. Intervention %s décalée du %s au %s pour la ressource %s.',
    {3}'Intervention décalée',
    {4}'Intervention non générée',
    {5}'Affaire : %s. Tiers : %s. Aucune intervention n''a été générée pour la tâche %s. ',
    {6}'Affaire : %s. Aucune intervention n''a été générée pour les tâches en heures.',
    {7}'Si vous avez généré en utilisant ''Ajouter au planning existant'', vous planifiez à partir de la dernière date de génération précisée dans l''onglet ''Règles de la fiche Tâche''. ',
    {8}'Attention à la date de fin de génération paramètrée qui est actuellement au %s.'
    );

//
// Fonctions générales liées à la ressource
//
(******************************************************************************)
// RessourceDuSalarie
// Entrée : Code salarié pour lequel on cherche la ressource associée
//
// Sortie : code ressource s'il existe et '' sinon.
(******************************************************************************)

function RessourceDuSalarie(AcsCodeSalarie: string): string;
var
  sSelectR: string;
  QSelectR: TQuery;
begin
  Result := '';
  QSelectR := nil;
  sSelectR := 'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_SALARIE="' + AcsCodeSalarie + '"';

  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);
    if not QSelectR.EOF then
      // Si la sélection n'est pas vide
    begin
      Result := QSelectR.FindField('ARS_RESSOURCE').AsString;
    end;
  finally
    Ferme(QSelectR);
  end;
end;

//
// OBJET TAFO_RESSOURCES
//

destructor TAFO_Ressources.Destroy;
var
  i: integer;
begin
  for i := 0 to count - 1 do
  begin
    Objects[i].Free;
  end;

  inherited;
end;

(******************************************************************************)
// RessourcesD1Competence : donne la liste des ressources possedant une competence
//
// entrée : code fonction recherché
(******************************************************************************)

procedure TAFO_Ressources.RessourcesD1Competence(AcsCompetence: string);
var
  sSelectR: string;
  QSelectR: TQuery;
  vrRessource: TAFO_Ressource;
  TobCompetRess: TOB;
  i: integer;
begin
  QSelectR := nil;

  sSelectR := 'SELECT DISTINCT ACR_RESSOURCE FROM COMPETRESSOURCE WHERE ACR_COMPETENCE="' + AcsCompetence + '"';
  TobCompetRess := TOB.Create('', nil, -1);
  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);

    TobCompetRess.LoadDetailDB('', '', '', QSelectR, true);
    for i := 0 to TobCompetRess.Detail.count - 1 do
    begin
      vrRessource := TAFO_Ressource.Create(string(TobCompetRess.Detail[i].GetValue('ACR_RESSOURCE')));
      if (vrRessource <> nil) then
        if (vrRessource.tob_Champs <> nil) then
          AddObject(TobCompetRess.Detail[i].GetValue('ACR_RESSOURCE'), vrRessource);
    end;

  finally
    Ferme(QSelectR);
    TobCompetRess.Free;
  end;
end;

(******************************************************************************)
// RessourceD1User : donne le code ressource du User en fourni en entrée
//
// entrée : code User recherché
(******************************************************************************)

function TAFO_Ressources.RessourceD1User(AcsUser: string): string;
var
  sSelectR: string;
  QSelectR: TQuery;
begin
  Result := '';
  QSelectR := nil;
  sSelectR := 'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE="' + AcsUser + '"';

  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);
    if not QSelectR.EOF then
      // Si la sélection n'est pas vide
    begin
      Result := QSelectR.FindField('ARS_RESSOURCE').AsString;
    end;
  finally
    Ferme(QSelectR);
  end;
end;

// PL le 08/02/06 : Ajout de la ressource trouvée dans la liste
function TAFO_Ressources.AddRessourceD1User(AcsUser: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false): integer;
var
  sRessCode: string;
begin
  Result := -1;

	// Si le user demandé est le user courant, on a déjà la ressource sous la main
  // sauf au chargement de l'appli ou en dehors du contexte GC, où elle est vide
{$IFDEF GCGC}
  if (AcsUser = V_PGI.User) and (VH_GC.RessourceUser <> '') then
		sRessCode := VH_GC.RessourceUser
	else
{$ENDIF GCGC}
		sRessCode := RessourceD1User(V_PGI.User);

	if (sRessCode <> '') then
		Result := AddRessource(sRessCode, AcbAvecCalendrier, AcbLeger, AcbAvecLibres);

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/10/2002
Modifié le ... :   /  /
Description .. : donne la liste des ressources affectées à une tâche d'une
                 affaire
                 remarque : cette liste ne ramene que les ressources actives
Mots clefs ... :
*****************************************************************}
{$IFDEF AFFAIRE}

procedure TAFO_Ressources.RessourceD1TacheAffaire(AcsAffaire, AcsNumeroTache: string);
var
  sSelectR: string;
  QSelectR: TQuery;
  vrRessource: TAFO_Ressource;
  TobRessources: TOB;
  i: Integer;

begin
  QSelectR := nil;

  sSelectR := 'SELECT ATR_RESSOURCE FROM TACHERESSOURCE ';
  sSelectR := sSelectR + 'WHERE ATR_AFFAIRE="' + AcsAffaire + '" ';
  sSelectR := sSelectR + ' AND ATR_STATUTRES = "ACT"';
  if AcsNumeroTache <> '' then
    sSelectR := sSelectR + ' AND ATR_NUMEROTACHE=' + AcsNumeroTache;

  TobRessources := TOB.Create('', nil, -1);
  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);

    TobRessources.LoadDetailDB('', '', '', QSelectR, true);
    for i := 0 to TobRessources.Detail.count - 1 do
    begin
      vrRessource := TAFO_Ressource.Create(string(TobRessources.Detail[i].GetValue('ATR_RESSOURCE')));
      if (vrRessource <> nil) then
        if (vrRessource.tob_Champs <> nil) then
          AddObject(TobRessources.Detail[i].GetValue('ATR_RESSOURCE'), vrRessource);
    end;
  finally
    Ferme(QSelectR);
    TobRessources.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 13/09/2004
Modifié le ... :   /  /
Description .. : donne la TOB des temps d'absence des ressources présentes dans la liste TobRessource
                  dans le même ordre d'indice, sur l'intervalle des dates en entrée
                  avec 3 colonnes en sortie :
                  ARS_RESSOURCE : code de la ressource
                  PERIODE       : période associée (du type AAAAMM pour la périodicité MOIS et AAAASS pour la périodicité SEM)
                  NBJABSENCE    : le nombre de jours d'absence
                 REMARQUE IMPORTANTE : la libération de la TOB est à la charge de l'appelant
Mots clefs ... :
*****************************************************************}

function TAFO_Ressources.TOBAbsencesDesRessources(pDtDebut, pDtFin: TDateTime; pPeriodicite: string; TobRessource: TOB = nil): TOB;
var
  ii: integer;
  TOBRess: TOB;
  bChargeLesAbsences: boolean;
begin
  Result := TOB.Create('Les absences', nil, -1);

  for ii := 0 to count - 1 do
    // On boucle sur toutes les ressources de la liste
  begin
    bChargeLesAbsences := true;
    if (TobRessource <> nil) then
      // S'il y a une TOB de ressources en entrée, on ne charge que les absences de cette tob
      // sinon, on les charge toutes
    begin
      bChargeLesAbsences := false;
      TOBRess := TobRessource.FindFirst(['ARS_RESSOURCE'], [strings[ii]], false);
      if (TOBRess <> nil) then
        bChargeLesAbsences := true;
    end;

    if bChargeLesAbsences then
      // Charge les absences de la ressource courante et les range dans la TOB Result
      // qui devra être désallouée par l'appelant
      TAFO_Ressource(Objects[ii]).ChargeTOBAbsenceDeLaRessource(Result, pDtDebut, pDtFin, pPeriodicite);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 28/10/2004
Modifié le ... :   /  /
Description .. : donne la TOB des temps de disponibilité des ressources présentes dans la liste TobRessource
                  dans le même ordre d'indice, sur l'intervalle des dates en entrée
                  avec 3 colonnes en sortie :
                  ARS_RESSOURCE : code de la ressource
                  PERIODE       : période associée (du type AAAAMM pour la périodicité MOIS et AAAASS pour la périodicité SEM)
                  NBHDISPO      : le nombre d'heures de disponibilité
                 REMARQUE IMPORTANTE : la libération de la TOB est à la charge de l'appelant
Mots clefs ... :
*****************************************************************}

function TAFO_Ressources.TOBDisponibiliteDesRessources(pDtDebut, pDtFin: TDateTime; pPeriodicite: string; TobRessource: TOB = nil): TOB;
var
  ii: integer;
  TOBRess: TOB;
  bChargeLesDispos: boolean;
begin
  Result := TOB.Create('Les Dispos', nil, -1);

  for ii := 0 to count - 1 do
    // On boucle sur toutes les ressources de la liste
  begin
    bChargeLesDispos := true;
    if (TobRessource <> nil) then
      // S'il y a une TOB de ressources en entrée, on ne charge que les dispos de cette tob
      // sinon, on les charge toutes
    begin
      bChargeLesDispos := false;
      TOBRess := TobRessource.FindFirst(['ARS_RESSOURCE'], [strings[ii]], false);
      if (TOBRess <> nil) then
        bChargeLesDispos := true;
    end;

    if bChargeLesDispos then
      // Charge les dispos de la ressource courante et les range dans la TOB Result
      // qui devra être désallouée par l'appelant
      TAFO_Ressource(Objects[ii]).ChargeTOBDispoDeLaRessource(Result, pDtDebut, pDtFin, pPeriodicite);
  end;
end;

{$ENDIF}

(******************************************************************************)
// RessourcesD1Fonction : donne la liste des ressources ayant eu ou ayant actuellement
//                         une fonction. Pour savoir si une des fonctions est valable
//                         à une date donnée, il faut appeler le service 'FonctionDeLaRessource'
//                         ressource par ressource.
//
// entrée : code fonction recherché
(******************************************************************************)

procedure TAFO_Ressources.RessourcesD1Fonction(AcsFonction: string);
var
  sSelectR: string;
  QSelectR: TQuery;
  vrRessource: TAFO_Ressource;
  TobRessources: TOB;
  i: integer;
begin
  QSelectR := nil;
  sSelectR := 'SELECT DISTINCT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_FONCTION1="' + AcsFonction + '" OR ARS_FONCTION2="' + AcsFonction + '" OR ARS_FONCTION3="' + AcsFonction + '" OR ARS_FONCTION4="' + AcsFonction + '"';

  TobRessources := TOB.Create('', nil, -1);
  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);

    TobRessources.LoadDetailDB('', '', '', QSelectR, true);
    for i := 0 to TobRessources.Detail.count - 1 do
    begin
      vrRessource := TAFO_Ressource.Create(string(TobRessources.Detail[i].GetValue('ARS_RESSOURCE')));
      if (vrRessource <> nil) then
        if (vrRessource.tob_Champs <> nil) then
          AddObject(TobRessources.Detail[i].GetValue('ARS_RESSOURCE'), vrRessource);
    end;

  finally
    Ferme(QSelectR);
    TobRessources.Free;
  end;
end;

function TAFO_Ressources.AddRessource(AcsRessource: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false): integer;
var
  vrRessource: TAFO_Ressource;
begin
  vrRessource := nil;
  if (AcsRessource = '') then
  begin
    Result := -2;
    exit;
  end;

  Result := IndexOf(AcsRessource);
  if (Result = -1) then
    // le code n'est pas encore dans la liste
  begin
    try
      vrRessource := TAFO_Ressource.Create(AcsRessource, AcbAvecCalendrier, AcbLeger, AcbAvecLibres);
      if (vrRessource <> nil) then
        if (vrRessource.tob_Champs <> nil) then
          Result := AddObject(AcsRessource, vrRessource)
        else
          vrRessource.Free;
    except
      // En cas d'erreur, on désalloue ce qu'on a alloué
      vrRessource.Free;
    end;
  end
  else
  begin
    // Si on a deja chargé la ressource, que l'on devrait avoir chargé aussi son calendrier
    // mais que le calendrier est vide, on va le charger
    if (AcbAvecCalendrier = true) and (TAFO_Ressource(GetRessource(Result)).TobCalendrier = nil) then
      TAFO_Ressource(GetRessource(Result)).ChargeCalendriers(true);

    // Si on a déjà chargé la ressource, que l'on devrait avoir tous les champs
    // mais qu'elle est seulement en chargement léger, on va charger le reste
//    if (AcbLeger = true) and
//				((TAFO_Ressource(GetRessource(Result)).gbCreateLeger = false) then
		// PL le 08/02/06 : inversion du test qui était faux
    if (AcbLeger = false) and
				(TAFO_Ressource(GetRessource(Result)).gbCreateLeger = true) then
      TAFO_Ressource(GetRessource(Result)).ChargeTousLesChamps('');

  end;
end;

function TAFO_Ressources.TOBRessources(AcsTitreTOB: string): TOB;
var
  i: integer;
  TOBFille: TOB;
begin
  Result := TOB.Create(AcsTitreTOB, nil, -1);

  for i := 0 to count - 1 do
  begin
    TOBFille := TOB.Create('', Result, -1);
    TOBFille.Dupliquer(TAFO_Ressource(Objects[i]).tob_Champs, TRUE, TRUE, TRUE);

    //   TAFO_Ressource(Objects[i]).tob_Champs.ChangeParent(Result, -1);
  end;
end;

function TAFO_Ressources.GetRessource(Index: Integer): TAFO_Ressource;
begin
  result := TAFO_Ressource(Objects[Index]);
end;

//
//
// OBJET TAFO_Ressource
//
//
{$IFDEF PLUGIN}
constructor TAFO_Ressource.Create(AcsCodeRessource: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false;APlugIn :TVignettePlugIn=nil);
var sSQL : string;
{$ELSE}
constructor TAFO_Ressource.Create(AcsCodeRessource: string; AcbAvecCalendrier: boolean = false; AcbLeger: boolean = false; AcbAvecLibres: boolean = false);
var
  bRessTrouvee: boolean;
  QQ: TQuery;
  sSQL: String;
{$ENDIF !PLUGIN}
begin
  gbCreateLeger := AcbLeger;
  gbCreateAvecLibres := AcbAvecLibres;
  // PL le 08/02/06 : Ajout des champs libres optionnellement
  sSQL := 'SELECT ARS_RESSOURCE,ARS_SALARIE,ARS_TYPERESSOURCE,ARS_LIBELLE,ARS_UNITETEMPS,ARS_STANDCALEN,ARS_CALENSPECIF,ARS_ARTICLE,ARS_ETABLISSEMENT,ARS_DEPARTEMENT';
  if AcbAvecLibres then
          sSQL := sSQL + ',ARS_LIBRERES1,ARS_LIBRERES2,ARS_LIBRERES3,ARS_LIBRERES4,ARS_LIBRERES5,ARS_LIBRERES6,ARS_LIBRERES7,ARS_LIBRERES8,ARS_LIBRERES9,ARS_LIBRERESA,ARS_EQUIPERESS';
  sSQL := sSQL + ' FROM RESSOURCE WHERE ARS_RESSOURCE="' + AcsCodeRessource + '"';

{$IFDEF PLUGIN}
   PlugIn := APlugIn;
   tob_Champs := PlugIn.OpenSelectInCache(sSQL); //AB-200609-
   if (tob_Champs <> nil) and (tob_Champs.detail.count > 0) then
   begin
     tob_Champs := tob_Champs.detail[0];
   end;
{$ELSE}
  bRessTrouvee := false;
  Tob_Champs := Tob.Create('RESSOURCE', nil, -1);
  if AcbLeger then // PL le 02/11/04 : create léger
  begin
    try
      QQ := OpenSQL(sSQL , true,-1,'',true);
      if not QQ.EOF then
      begin
        bRessTrouvee := true;
        tob_Champs.selectDB('', QQ);
      end;
    finally
      Ferme(QQ);
    end;
  end
  else
    if tob_Champs.SelectDB('"' + AcsCodeRessource + '"', nil) then
  begin
    bRessTrouvee := true;
  end;
  if not bRessTrouvee then
  begin
    tob_Champs.Free;
    tob_Champs := nil;
  end;
{$ENDIF !PLUGIN}

  if (tob_Champs <> nil) then
  begin
    gsUniteValo := tob_Champs.GetValue('ARS_UNITETEMPS');
    gbPVHT := true;
    gDatePrix := 0;
    gdTempsAnnuel := 0;

    // Chargement des calendriers de la ressource
    if AcbAvecCalendrier then
      ChargeCalendriers(true);

    // PL le 04/10/05 : On charge de toutes façon les jours fériés
//{$IFDEF AFFAIRE}
    gobj_JourFerie := TJourFerie.Create;
//{$ENDIF}            //AB-200607-
    gsUniteActivite := GetParamSocSecur('SO_AFMESUREACTIVITE', '') ;

  end;
end;

(******************************************************************************)
// ChargeTousLesChamps : Charge tous les champs de la ressource en entrée
//                       si pas de ressource en entrée, considère que c'est un rechargement
//
// entrée :
(******************************************************************************)

function TAFO_Ressource.ChargeTousLesChamps(AcsRessource: string): boolean;
var
  sCodeRessource: string;
begin
  Result := false;
  if (AcsRessource = '') then
  begin
    // si on ne précise pas la ressource en entrée, ça ne peut être qu'un rechargement
    if (Tob_Champs = nil) then
      exit; // S'il n'y avait rien avant, on sort

    sCodeRessource := Tob_Champs.GetValue('ARS_RESSOURCE');
  end
  else
    sCodeRessource := AcsRessource;

  if (Tob_Champs <> nil) then
    Tob_Champs.Free;

  Tob_Champs := Tob.Create('RESSOURCE', nil, -1);

  if not Tob_Champs.SelectDB('"' + sCodeRessource + '"', nil) then
  begin
    Tob_Champs.Free;
    Tob_Champs := nil;
  end
  else
    Result := true;

end;

// Charge les temps de la ressource sur la période d'analyse 35H

(******************************************************************************)
// ChargeTempsAnnuel35H : Charge à l'instant t le temps de saisie annuel dans la variable
//                          prévue à cet effet : sur l'intervalle d'analyse 35h des paramètres
//
// entrée :
(******************************************************************************)

procedure TAFO_Ressource.ChargeTempsAnnuel35H;
begin
  gdTempsAnnuel := TempsAnnuelRessource('REA', GetParamSocSecur('SO_AFDATEDEB35', iDate1900), GetParamSocSecur('SO_AFDATEFIN35', iDate2099));
end;

(******************************************************************************)
// DetDerniereDateEffetCalendrier : donne la derniere date d'effet du calendrier règle juste inférieure
//                                  à la date en entrée et l'indice de la première TOB calendrier valable
//                                  à la date d'effet
//
// entrée : Date de référence, Indice de la première TOB calendrier
(******************************************************************************)

function TAFO_Ressource.DetDerniereDateEffetCalendrier(AcdDateRef: TDateTime; var AviIndice: integer): TDateTime;
var
  i: integer;
begin
  AviIndice := -1;
  Result := idate1900;

  // Ne charge les calendriers que s'il n'ont pas deja été chargés
  ChargeCalendriers(false);

  if (TobCalendrierRegle <> nil) then
  begin
    // La Tob étant triée en ordre décroissant, on récupère la première date d'effet juste inférieure
    // à la date référence
    i := 0;
    while (i < TobCalendrierRegle.Detail.count) and (TobCalendrierRegle.Detail[i].GetValue('ACG_DATEEFFET') > AcdDateRef) do
      Inc(i);

    if (i < TobCalendrierRegle.Detail.count) then
    begin
      Result := TobCalendrierRegle.Detail[i].GetValue('ACG_DATEEFFET');
      AviIndice := i;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 24/03/2003
Modifié le ... : 27/05/2004 par PL
Description .. : Recherche si la paie est utilisée : en fait, si des lignes d'absences ont été
                  saisies pour la ressource courante
Mots clefs ... : PAIE;
*****************************************************************}

function TAFO_Ressource.PaieUtilisee: Boolean;
var
  vSt: string;
  //  vQr : TQuery;

begin
  result := false;
  try
    try
      if (giAbsencePaie = 0) then
      begin
        vSt := 'SELECT PCN_SALARIE FROM ABSENCESALARIE WHERE PCN_SALARIE="' + tob_Champs.GetValue('ARS_SALARIE')
          + '" AND (PCN_TYPEMVT="ABS" OR PCN_TYPECONGE="PRI")'
          + ' AND PCN_MVTDUPLIQUE<>"X"	AND PCN_VALIDRESP="VAL" AND PCN_CODETAPE<>"X"'
          //C.B 13/04/2006
    			// ne pas ternir compte des absences annulées
    			+ ' AND PCN_ETATPOSTPAIE <> "NAN" ';

        if ExisteSQL(vSt) then
          result := true;
      end
      else
        if (giAbsencePaie = 1) then
        result := true
      else // cas du -1
        result := false;

    finally
      if result then
        giAbsencePaie := 1
      else
        giAbsencePaie := -1;
    end;
  except
    giAbsencePaie := 0;
  end;

  { PL le 27/05/04 :
  je l'ai réduit au salarié en cours puisqu'on est dans une ressource
  j'ai ajouté une variable giAbsencePaie remplie par cette fonction pour éviter de relancer n fois la même req

  vSt := 'SELECT COUNT(*) AS NB FROM ABSENCESALARIE ';
  try
    vQr := OpenSQL(vSt, True);
    if Not vQr.EOF then
    begin
      if vQr.FindField('NB').AsInteger = 0 then
        result := false
      else
        result := true;
    end;
  finally
    Ferme(vQr);
  end;}
end;

{$IFDEF AFFAIRE}
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/10/2002
Modifié le ... :   /  /
Description .. : genere la liste des jours à planifiés après application
                 des règles et des contraintes de la ressource
Mots clefs ... :
*****************************************************************}
function TAFO_Ressource.GenererPlanning(var pInNumLigne: Integer; pTob: Tob; pAFReglesTache: TAFRegles;
  pStAffaire: string; pTache: RecordTache; pArticle: RecordArticle;
  pBoValorisation, pBoAjout: Boolean; pPanel: ThPanel): Boolean;
var
  vAFReglesRes: TAFRegles;
  vListeJours: TListDate;
  vRdDuree: Double; // durée des interventions
  vInNbJoursDecal: Integer; // nombre de jours de decalage maximum
  vStMethodeDecal: string; // méthode de décalage des jours
  vRdQte: Double;
  vTobJoursOccupes: Tob;
  vSt : String;

  function QuantiteRessource(pRdQte: Double): Integer;
  begin
    result := 0;
    if trunc(vRdQte) <> vRdQte then
      result := result + 1
    else
      result := trunc(vRdQte);
  end;

begin
  vRdDuree := 1;
  vRdQte := 0;
  result := true;
  vInNbJoursDecal := 0;
  vStMethodeDecal := 'A';

  fSlRapportAnn := TStringList.Create;
  fSlRapportDec := TStringList.Create;
  vAFReglesRes := TAFRegles.create;
  try
    // regles par ressource
    // attention : ramene egalement la quantite initiale de la ressource (vRdQte) utilise dans le else
    if vAFReglesRes.LoadDBReglesRes(pStAffaire, pTache.StNumeroTache, tob_Champs.GetValue('ARS_RESSOURCE'), vRdQte, Ptache.boCompteur) then
    try

      // C.B 27/07/2005
      // si option pas de surbooking
      // Chargement des jours occupés dans le planning
      if not GetParamsocSecur('SO_AFSURBOOKING', False) then
        ;
      vTobJoursOccupes := LoadListeJoursOccupes(tob_Champs.GetValue('ARS_RESSOURCE'), pAFReglesTache.DtDateDeb, pAFReglesTache.DtDateFin);

      // calculer la quantité par ressource
      if pTache.BoCompteur then
        vListeJours := vAFReglesRes.GenereListeJours(vRdDuree, vInNbJoursDecal, vStMethodeDecal, vAFReglesRes.QuantiteRessource)
      else
        vListeJours := vAFReglesRes.GenereListeJours(vRdDuree, vInNbJoursDecal, vStMethodeDecal, cInNoCompteur);

      if vListeJours <> nil then
      begin
        try
          ControleLastDateGene(vListeJours, pTache.StLastDateGene, pBoAjout);
          ControleListeEtDecale(vListeJours, vRdDuree, vInNbJoursDecal, vStMethodeDecal,
            vTobJoursOccupes, pStAffaire,
            tob_Champs.GetValue('ARS_RESSOURCE'), pTache);

          if vListeJours.count = 0 then
					begin
          	{5 'Affaire : %s. Tiers : %s. Aucune intervention n''a été générée pour la tâche %s. Si vous avez généré en utilisant "Ajouter au planning existant", vous planifiez à partir de la dernière date de génération précisée dans l''onglet "Règles de la fiche Tâche".'}
            vSt := format(TraduitGa(TexteMessage[5]), [PStAffaire, pTache.StTiers, pTache.StLibTache]) + TraduitGa(TexteMessage[7]);
            if pAFReglesTache.DtDateFin < now then
              vSt := vSt + format(TraduitGa(TexteMessage[8]), [dateToStr(pAFReglesTache.DtDateFin)]);
//Modif FV   	AddJournal(fSlRapportAnn, vSt);
          end
          // C.B 02/11/2006
          // on controle si la ressource à bien un calendrier
          // sinon, on ne peut ps générer
          {else if not AUnCalendrier then
          begin
            vSt := format(TraduitGa(TexteMessage[9]), [tob_Champs.GetValue('ARS_RESSOURCE')]);
          	AddJournal(fSlRapportAnn, vSt);
          end} 
          else
            result := GenererLignesPlanning(vListeJours, pInNumLigne, pTob,
              pStAffaire, pTache, pArticle, pBoValorisation, vRdDuree);
        finally
    			{3 'Intervention décalée',}
//Modif FV CreeJournaux(fSlRapportDec, TexteMessage[3], 'AVE', 'PLG');
			    {4 'Intervention non générée',}
//Modif FV CreeJournaux(fSlRapportAnn, TexteMessage[4], 'AVE', 'PLG');
        end;
      end;
    finally
      vListeJours.Free;
    end
      // regles par taches
    else
    try
      // si option pas de surbooking
      // Chargement des jours occupés dans le planning
      if not GetParamsocSecur('SO_AFSURBOOKING', False) then
        ;
      vTobJoursOccupes := LoadListeJoursOccupes(tob_Champs.GetValue('ARS_RESSOURCE'), pAFReglesTache.DtDateDeb, pAFReglesTache.DtDateFin);

      if pTache.BoCompteur then
        // calculer la quantité par ressource
        vListeJours := pAFReglesTache.GenereListeJours(vRdDuree, vInNbJoursDecal, vStMethodeDecal, QuantiteRessource(vRdQte))
      else
        vListeJours := pAFReglesTache.GenereListeJours(vRdDuree, vInNbJoursDecal, vStMethodeDecal, cInNoCompteur);

      if vListeJours <> nil then
      begin
        try
          ControleLastDateGene(vListeJours, pTache.StLastDateGene, pBoAjout); //pLastDateGene
          ControleListeEtDecale(vListeJours, vRdDuree, vInNbJoursDecal, vStMethodeDecal,
            vTobJoursOccupes, pStAffaire,
            tob_Champs.GetValue('ARS_RESSOURCE'), pTache);
          if vListeJours.count = 0 then
					begin
          	{5 'Affaire : %s. Tiers : %s. Aucune intervention n''a été générée pour la tâche %s. Si vous avez généré en utilisant "Ajouter au planning existant", vous planifiez à partir de la dernière date de génération précisée dans l''onglet "Règles de la fiche Tâche".'}
            vSt := format(TraduitGa(TexteMessage[5]), [PStAffaire, pTache.StTiers, pTache.StLibTache]) + TraduitGa(TexteMessage[7]);
            if pAFReglesTache.DtDateFin < now then
              vSt := vSt + format(TraduitGa(TexteMessage[8]), [dateToStr(pAFReglesTache.DtDateFin)]);
//Modif FV 	AddJournal(fSlRapportAnn, vSt);
					end

          // C.B 02/11/2006
          // on controle si la ressource à bien un calendrier
          // sinon, on ne peut ps générer
{          else if not AUnCalendrier then
          begin
            vSt := format(TraduitGa(TexteMessage[9]), [tob_Champs.GetValue('ARS_RESSOURCE')]);
          	AddJournal(fSlRapportAnn, vSt);
          end}
          else
            result := GenererLignesPlanning(vListeJours, pInNumLigne, pTob,
              pStAffaire, pTache, pArticle, pBoValorisation, vRdDuree);
        finally
			    {3'Intervention décalée',}
//Modif FV CreeJournaux(fSlRapportDec, TexteMessage[3], 'AVE', 'PLG');
			    {4'Intervention non générée',}
//Modif FV CreeJournaux(fSlRapportAnn, TexteMessage[4], 'AVE', 'PLG');
        end;
      end;
    finally
      vListeJours.Free;
      FreeAndNil(vTobJoursOccupes);
    end;
  finally
    vAFReglesRes.Free;
    fSlRapportAnn.Free;
    fSlRapportDec.Free;
  end;
end;
 
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.GenererLignesPlanning(pListeJours: TListDate; var pInNumLigne: Integer; pTob: Tob;
  pStAffaire: string; pTache: RecordTache; pArticle: RecordArticle;
  pBoValorisation: Boolean; pRdDuree: Double): Boolean;
var
  vTob: Tob;
  i: Integer;
  vInNumLigne: Integer;
  vStAff0 : string;     
  vStAff1 : string;
  vStAff2 : string;
  vStAff3 : string;
  vStAvenant : string;

  vDtFin          : TDateTime;
  vRdDureeEnJour  : Double;
  vTOBAffaires    : Tob;
  vAFOAssistants  : TAFO_Ressources;
  vTobArticles    : Tob;
//  vBoTacheHeure   : Boolean;

begin

  result := true;
//  vBoTacheHeure := False;
  vInNumLigne := 0;
  CodeAffaireDecoupe(pStAffaire, vStAff0, vStAff1, vStAff2, vStAff3, vStAvenant, taModif, false);
  vAFOAssistants := TAFO_Ressources.Create;
  vTOBAffaires := TOB.Create('Les Affaires', nil, -1);
  vTOBArticles := TOB.Create('Les Articles', nil, -1);

  try
    //    AFRemplirTOBAffaire('', nil, TobAffaires);
    //    RemplirTOBAssistant( nil, vAFOAssistants, '');

    for i := 0 to pListeJours.count - 1 do
    begin
      //C.B 21/03/2006
      // on ne genere que les taches en jours
      //vSt := 'SELECT ATA_AFFAIRE FROM TACHE WHERE ATA_UNITETEMPS = "H" AND ATA_AFFAIRE = "' + pStAffaire + '"' + ' AND ATA_NUMEROTACHE = ' + pTache.StNumeroTache;
      //if not ExisteSql(vSt) then
      //begin
        MoveCurProgressForm('');
        vTob := TOB.Create('AFPLANNING', pTob, -1);

        vInNumLigne := pInNumLigne + i;
        vTob.PutValue('APL_AFFAIRE', pStAffaire);
        vTob.PutValue('APL_NUMEROLIGNE', vInNumLigne);
        vTob.PutValue('APL_TYPELIGNEPLA', 'TAC');
        vTob.PutValue('APL_TYPEPLA', 'PLA');
        vTob.PutValue('APL_NUMEROTACHE', pTache.StNumeroTache);
        //mcd 26/07/04 ajout trait zone libres
        vTob.PutValue('APL_LIBRETACHE1', pTache.StLib1);
        vTob.PutValue('APL_LIBRETACHE2', pTache.StLib2);
        vTob.PutValue('APL_LIBRETACHE3', pTache.StLib3);
        vTob.PutValue('APL_LIBRETACHE4', pTache.StLib4);
        vTob.PutValue('APL_LIBRETACHE5', pTache.StLib5);
        vTob.PutValue('APL_LIBRETACHE6', pTache.StLib6);
        vTob.PutValue('APL_LIBRETACHE7', pTache.StLib7);
        vTob.PutValue('APL_LIBRETACHE8', pTache.StLib8);
        vTob.PutValue('APL_LIBRETACHE9', pTache.StLib9);
        vTob.PutValue('APL_LIBRETACHEA', pTache.StLibA);
        vTob.PutValue('APL_BOOLLIBRE1', pTache.StBool1);
        vTob.PutValue('APL_BOOLLIBRE2', pTache.StBool2);
        vTob.PutValue('APL_BOOLLIBRE3', pTache.StBool3);
        vTob.PutValue('APL_ValLibre1', pTache.RdVal1);
        vTob.PutValue('APL_ValLibre2', pTache.RdVal2);
        vTob.PutValue('APL_ValLibre3', pTache.RdVal3);
        vTob.PutValue('APL_CharLibre1', pTache.StChar1);
        vTob.PutValue('APL_CharLibre2', pTache.StChar2);
        vTob.PutValue('APL_CharLibre3', pTache.StChar3);
        vTob.PutValue('APL_DateLibre1', pTache.DtDate1);
        vTob.PutValue('APL_DateLibre2', pTache.DtDate2);
        vTob.PutValue('APL_DateLibre3', pTache.DtDate3);
        vTob.PutValue('APL_RESSOURCE', tob_Champs.GetValue('ARS_RESSOURCE'));
        if pTache.StFctTache <> '' then
          vTob.PutValue('APL_FONCTION', pTache.StFctTache)
        else
          vTob.PutValue('APL_FONCTION', '');
        vTob.PutValue('APL_ARTICLE', pTache.StArticle);
        vTob.PutValue('APL_TYPEARTICLE', pArticle.StTypeArticle);
        vTob.PutValue('APL_CODEARTICLE', pArticle.StCodeArticle);

        // recuperer le mode de facturation de l'article
        vTob.PutValue('APL_ACTIVITEREPRIS', pTache.StActiviteRepris);
        vTob.PutValue('APL_TIERS', pTache.StTiers);
        vTob.PutValue('APL_ETATLIGNE', getparamsocSecur('SO_AFPLANNINGETAT', ''));
        vTob.PutValue('APL_LIGNEGENEREE', 'X');
        vTob.PutValue('APL_ACTIVITEGENERE', '-');
        vTob.PutValue('APL_AFFAIRE0', vStAff0);
        vTob.PutValue('APL_AFFAIRE1', vStAff1);
        vTob.PutValue('APL_AFFAIRE2', vStAff2);
        vTob.PutValue('APL_AFFAIRE3', vStAff3);
        vTob.PutValue('APL_AVENANT', vStAvenant);
        vTob.PutValue('APL_ACTIVITEEFFECT', '-');

        // on genere les items du planning dans la meme unité de temps
        // que la tache
        vTob.PutValue('APL_UNITETEMPS', pTache.StUnite);
        vTob.PutValue('APL_QTEPLANIFIEE', pRdDuree);
        vTob.PutValue('APL_QTEPLANIFUREF', AFConversionUnite(pTache.StUnite, VH_GC.AFMESUREACTIVITE, pRdDuree));
        vTob.PutValue('APL_INITPTPR', '0');
        vTob.PutValue('APL_REALPTPR', '0');
        vTob.PutValue('APL_LIBELLEPLA', pTache.StLibTache);

        vTob.PutValue('APL_DATEDEBPLA', pListeJours.Items[i]);

        // date fin = date deb + qte
        // regle utilisée : si on ne tombe pas sur un nombre de jour entier
        // on arrondi au nombre de jour supérieur
        vRdDureeEnJour := AFConversionUnite(pTache.StUnite, 'J', pRdDuree);
        if int(vRdDureeEnJour) < pRdDuree then
          vRdDureeEnJour := int(vRdDureeEnJour) + 1;
        vDtFin := plusDate(pListeJours.Items[i], round(vRdDureeEnJour - 1), 'J');
        vTob.PutValue('APL_DATEFINPLA', vDtFin);

        // C.B 29/06/2005
        // ajout de la gestion des heures
        vTob.putValue('APL_HEUREDEB_PLA', vTob.GetValue('APL_DATEDEBPLA') + getparamsocSecur('SO_AFAMDEBUT', 0));
        vTob.putValue('APL_HEUREFIN_PLA', vTob.GetValue('APL_DATEFINPLA') + getparamsocSecur('SO_AFPMDEBUT', 0));

        if getParamSocSecur('SO_AFALIGNREALISE', False) then
        begin
          vTob.PutValue('APL_QTEREALISE', vTob.GetValue('APL_QTEPLANIFIEE'));
          vTob.PutValue('APL_QTEREALUREF', vTob.GetValue('APL_QTEPLANIFUREF'));
        end;

        // on aligne les dates de realisation sur les dates de planification
        vTob.PutValue('APL_DATEDEBREAL', vTob.GetValue('APL_DATEDEBPLA'));
        vTob.PutValue('APL_DATEFINREAL', vTob.GetValue('APL_DATEFINPLA'));

        // C.B 29/06/2005
        // ajout de la gestion des heures
        vTob.putValue('APL_HEUREDEB_REAL', vTob.GetValue('APL_HEUREDEB_PLA'));
        vTob.putValue('APL_HEUREFIN_REAL', vTob.GetValue('APL_HEUREFIN_PLA'));

        // valorisation des lignes de planning
        if pBoValorisation then
          ValorisationPlanning (vTob, 'APL', vAFOAssistants, vTOBAffaires, vTobArticles, true, True);
        // on conserve la derniere date générée
        if i = (pListeJours.count - 1) then
          ExecuteSql('UPDATE TACHE SET ATA_LASTDATEGENE = "' + UsDateTime(pListeJours.Items[i]) + '" WHERE ATA_AFFAIRE = "' + pStAffaire + '" AND ATA_NUMEROTACHE = ' + pTache.StNumeroTache);
      end;
      //else
      //  vBoTacheHeure := True;
    	//end;

    if pListeJours.count <> 0 then
      pInNumLigne := vInNumLigne + 1;

    //if vBoTacheHeure then
    //  AddJournal(fSlRapportAnn, format(TraduitGa(TexteMessage[6]), [PStAffaire]));

  finally
    FreeAndNil(vTOBAffaires);
    FreeAndNil(vTobArticles);
    vAFOAssistants.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 13/09/2004
Modifié le ... :
Description .. : Compte le nombre de jours d'absence dans l'intervalle de date en entrée
                  et renvoie une tob avec les trois champs suivants :
                  ARS_RESSOURCE : ressource courante
                  PERIODE : periode associée (clé pour retrouver les absences)
                  NBJABSENCE : nb de jour d'absence pour cette periode

                 remarques :
                 - IMPORTANT : Met à jour la tob TobAbsencesParPeriode qui contient toutes les absences
                 déjà répertoriées (pour éviter de lancer n fois la même recherche) pour la ressource courante
                 cette TOB générale est désallouée à la destruction de L'objet courant.
                 - si les jours feries sont travaillés, on ne controle que le calendrier
                 - Les deux dates en entree sont inclues
                 - force le recalcul si bForcer = true
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.NbJoursAbsence(pDtDebut, pDtFin: TDateTime; sPeriode: string; bForcer: boolean = false): TOB;
var
  dNbJAbsence: double;
  pDtCourante: TDateTime;
  Ress: string;
  NbHDemiJournee, NbHeureJour: double;
  TobFille: TOB;
begin
  Result := nil;
  NbHDemiJournee := AFConversionUnite('J', 'H', 0.5);
  Ress := tob_Champs.GetValue('ARS_RESSOURCE');
  dNbJAbsence := 0;
  if (pDtDebut >= pDtFin) then
    exit;

  TobFille := nil;
  try
    // Si on ne force pas le recalcul, on regarde si la période en question est déjà dispo dans la TOB TobAbsencesParPeriode
    if (not bForcer) and (TobAbsencesParPeriode <> nil) then
      TobFille := TobAbsencesParPeriode.FindFirst(['ARS_RESSOURCE', 'PERIODE'], [Ress, sPeriode], false);

    if (TobFille = nil) then
      // Si la période en question n'est pas calculée ou si on force le recalcul,
      // calcul des absences de la période
    begin
      pDtCourante := pDtDebut;
      repeat
        JourTravaille(pDtCourante, NbHeureJour, false);

        if (NbHeureJour = 0) then
          dNbJAbsence := dNbJAbsence + 1
        else
          if (NbHeureJour = NbHDemiJournee) then
          dNbJAbsence := dNbJAbsence + 0.5;

        pDtCourante := pDtCourante + 1;
      until (pDtCourante > pDtFin);

      // On attache la tob resultat à la tob générale TobAbsencesParPeriode
      // qui contient toutes les absences déjà calculées de la ressource
      TobFille := TOB.Create('La Fille', TobAbsencesParPeriode, -1);
      TobFille.AddChampSupValeur('ARS_RESSOURCE', Ress);
      TobFille.AddChampSupValeur('PERIODE', sPeriode);
      TobFille.AddChampSupValeur('NBJABSENCE', dNbJAbsence);
    end;

  finally
    Result := TobFille;
  end;
end;
{$ENDIF}



{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 28/10/2004
Modifié le ... :
Description .. : Compte le nombre d'heures de disponibilité dans l'intervalle de date en entrée
                  et renvoie une tob avec les 4 champs suivants :
                  ARS_RESSOURCE : ressource courante
                  PERIODE : periode associée (clé pour retrouver les absences)
                  NBHDISPO : nb d'heures de dispo pour cette periode

                 remarques :
                 - IMPORTANT : Met à jour la tob TobDispoParPeriode qui contient toutes les Dispos
                 déjà répertoriées (pour éviter de lancer n fois la même recherche) pour la ressource courante
                 cette TOB générale est désallouée à la destruction de L'objet courant.
                 - si les jours feries sont travaillés, on ne controle que le calendrier
                 - Les deux dates en entree sont inclues
                 - force le recalcul si bForcer = true
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.NbHeuresDispo(pDtDebut, pDtFin: TDateTime; sPeriode: string; bForcer: boolean = false): TOB;
var
  pDtCourante: TDateTime;
  Ress: string;
  TobFille: TOB;
  NbHDispo, NbHeureDuJour: double;
begin
  Result := nil;
  Ress := tob_Champs.GetValue('ARS_RESSOURCE');
  if (pDtDebut >= pDtFin) then
    exit;

  TobFille := nil;
  try
    // Si on ne force pas le recalcul, on regarde si la période en question est déjà dispo dans la TOB TobDisposParPeriode
    if (not bForcer) and (TobDisposParPeriode <> nil) then
      TobFille := TobDisposParPeriode.FindFirst(['ARS_RESSOURCE', 'PERIODE'], [Ress, sPeriode], false);

    if (TobFille = nil) then
      // Si la période en question n'est pas calculée ou si on force le recalcul,
      // calcul des dispos de la période
    begin
      pDtCourante := pDtDebut;
      NbHDispo := 0;
      repeat
        // Si le jour est travaille on additionne le nombre d'heures
        JourTravaille(pDtCourante, NbHeureDuJour, true);

        NbHDispo := NbHDispo + NbHeureDuJour;

        pDtCourante := pDtCourante + 1;
      until (pDtCourante > pDtFin);

      // On attache la tob resultat à la tob générale TobDisposParPeriode
      // qui contient toutes les dispos déjà calculées de la ressource
      TobFille := TOB.Create('La Fille', TobDisposParPeriode, -1);
      TobFille.AddChampSupValeur('ARS_RESSOURCE', Ress);
      TobFille.AddChampSupValeur('PERIODE', sPeriode);
      TobFille.AddChampSupValeur('NBHDISPO', NbHDispo);
    end;

  finally
    Result := TobFille;
  end;
end;

{$IFDEF AFFAIRE}
{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 13/09/2004
Modifié le ... :   /  /
Description .. : donne la TOB des temps d'absence des ressources présentes dans la liste TobRessource ou
                toutes les ressources si TOBRessource est vide.
                TOBMere en entrée doit être allouée et reçoit en filles toutes les ressources et pour chacune
                les périodes assoiées à l'intervalle de dates et à la périodicité fournies en entrée

                  avec 3 colonnes en sortie :
                  ARS_RESSOURCE : code de la ressource
                  PERIODE       : période associée (du type AAAAMM pour la périodicité MOIS et AAAASS pour la périodicité SEM)
                  NBJABSENCE    : le nombre de jours d'absence

                 REMARQUE IMPORTANTE : la libération de la TOB est à la charge de l'appelant
Mots clefs ... :
*****************************************************************}

procedure TAFO_Ressource.ChargeTOBAbsenceDeLaRessource(TOBMere: TOB; pDtDebut, pDtFin: TDateTime; pPeriodicite: string);
var
  NumSem, NumAnnee: integer;
  sPeriodeCourante, sRessCourante: string;
  Annee, mois, jour: word;
  TobFille: TOB;
  DdebutPeriode, DfinPeriode: TDateTime;
  TobAbsenceRess: TOB;
  bForcer: boolean;
begin
  sRessCourante := tob_Champs.GetValue('ARS_RESSOURCE');

  // Init des champs date courante
  DfinPeriode := pDtDebut - 1;

  repeat // on scrute l'intervalle de dates fourni en entrée pour le décomposer en n période
    // suivant la périodicité en entrée

// nouvelles dates courantes
    DdebutPeriode := DfinPeriode + 1;
    DecodeDate(DdebutPeriode, Annee, mois, jour);
    if (pPeriodicite = 'MOIS') then
      DfinPeriode := FinDeMois(DdebutPeriode)
    else
      if (pPeriodicite = 'SEM') then
    begin
      NumSem := NumSemaine(DdebutPeriode, NumAnnee);
      DfinPeriode := PremierJourSemaine(NumSem, NumAnnee) + 7;
    end;

    if (DfinPeriode > pDtFin) then // on rectifie si on depasse la date de fin
      DfinPeriode := pDtFin;

    // On détermine la période courante
    if (pPeriodicite = 'MOIS') then
      sPeriodeCourante := inttostr(Annee) + inttostr(mois)
    else
      if (pPeriodicite = 'SEM') then
    begin
      NumSem := NumSemaine(DdebutPeriode, NumAnnee);
      sPeriodeCourante := inttostr(NumAnnee) + inttostr(NumSem);
    end;

    TobAbsenceRess := nil;
    if (TobCalendrier <> nil) or (TobAbsenceSalarie <> nil) then
    begin
      bForcer := false;
      if (pPeriodicite <> gsPeriodiciteCourante) then
        // Si la périodicité a changée, on force le recalcul
        bForcer := true;

      // Calcule le nombre de jours d'absence
      TobAbsenceRess := NbJoursAbsence(DdebutPeriode, DfinPeriode, sPeriodeCourante, bForcer);
    end;

    if (TobAbsenceRess <> nil) then
    begin
      gsPeriodiciteCourante := pPeriodicite;
      // On duplique la tob réponse pour la lier à la tob Mere qui est le résultat
      TobFille := TOB.Create('La Fille', TOBMere, -1);
      TobFille.Dupliquer(TobAbsenceRess, true, true);
    end;
  until (DfinPeriode >= pDtFin);
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 28/10/2004
Modifié le ... :   /  /
Description .. : donne la TOB des temps de dispo des ressources présentes dans la liste TobRessource ou
                toutes les ressources si TOBRessource est vide.
                TOBMere en entrée doit être allouée et reçoit en filles toutes les ressources et pour chacune
                les périodes assoiées à l'intervalle de dates et à la périodicité fournies en entrée

                  avec 3 colonnes en sortie :
                  ARS_RESSOURCE : code de la ressource
                  PERIODE       : période associée (du type AAAAMM pour la périodicité MOIS et AAAASS pour la périodicité SEM)
                  NBHDISPO      : le nombre d'heures de dispo

                 REMARQUE IMPORTANTE : la libération de la TOB est à la charge de l'appelant
Mots clefs ... :
*****************************************************************}

procedure TAFO_Ressource.ChargeTOBDispoDeLaRessource(TOBMere: TOB; pDtDebut, pDtFin: TDateTime; pPeriodicite: string);
var
  NumSem, NumAnnee: integer;
  sPeriodeCourante, sRessCourante: string;
  Annee, mois, jour: word;
  TobFille: TOB;
  DdebutPeriode, DfinPeriode: TDateTime;
  TobDispoRess: TOB;
  bForcer: boolean;
begin
  sRessCourante := tob_Champs.GetValue('ARS_RESSOURCE');

  // Init des champs date courante
  DfinPeriode := pDtDebut - 1;

  repeat // on scrute l'intervalle de dates fourni en entrée pour le décomposer en n période
    // suivant la périodicité en entrée

// nouvelles dates courantes
    DdebutPeriode := DfinPeriode + 1;
    DecodeDate(DdebutPeriode, Annee, mois, jour);
    if (pPeriodicite = 'MOIS') then
      DfinPeriode := FinDeMois(DdebutPeriode)
    else
      if (pPeriodicite = 'SEM') then
    begin
      NumSem := NumSemaine(DdebutPeriode, NumAnnee);
      //DfinPeriode := PremierJourSemaine (NumSem, NumAnnee) + 7;
      //AB-200509- si premier J semaine est lundi la semaine suivante on débutait le mardi
      DfinPeriode := PremierJourSemaine(NumSem, NumAnnee) + 6;
    end;

    if (DfinPeriode > pDtFin) then // on rectifie si on depasse la date de fin
      DfinPeriode := pDtFin;

    // On détermine la période courante
    if (pPeriodicite = 'MOIS') then
      sPeriodeCourante := inttostr(Annee) + inttostr(mois)
    else
      if (pPeriodicite = 'SEM') then
    begin
      NumSem := NumSemaine(DdebutPeriode, NumAnnee);
      sPeriodeCourante := inttostr(NumAnnee) + inttostr(NumSem);
    end;

    TobDispoRess := nil;
    if (TobCalendrier <> nil) or (TobAbsenceSalarie <> nil) then
    begin
      bForcer := false;
      if (pPeriodicite <> gsPeriodiciteCourante) then
        // Si la périodicité a changée, on force le recalcul
        bForcer := true;

      // Calcule le nombre d'heures de dispo
      TobDispoRess := NbHeuresDispo(DdebutPeriode, DfinPeriode, sPeriodeCourante, bForcer);
    end;

    if (TobDispoRess <> nil) then
    begin
      gsPeriodiciteCourante := pPeriodicite;
      // On duplique la tob réponse pour la lier à la tob Mere qui est le résultat
      TobFille := TOB.Create('La Fille', TOBMere, -1);
      TobFille.Dupliquer(TobDispoRess, true, true);
    end;
  until (DfinPeriode >= pDtFin);
end;


{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 13/11/2002
Modifié le ... : 10/09/04 par PL
Description .. : Verifie si le jour est travaillé, c'est a dire que le jour
                 est dans le calendrier et que ce n'est pas un jour férié ou de congés (rtt)
                 Retourne le nombre d'heures travaillées dans la journée
                 Un jour est travaille dès qu'une partie du temps de la journée est travaillée
                 c'est le nombre d'heures qui spécifie dans quelle proportion

                 remarque : si les jours feries sont travaillés,
                            on ne controle que le calendrier
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.JourTravaille(pDtWork: TDateTime; var NbHeure: double; bAxeHeures: boolean): Boolean;
var
  i: Integer;
  vInJourSemaine: Integer;
  vDtCourante: TDateTime;
  TobAbsenceDuJour: TOB;
  NbHAbsence, NbJAbsence, NbHeureTravMatin, NbHeureTravApMidi, NbHDemiJournee: double;
begin

  result := true;
  NbHeure := 0;
  NbHeureTravMatin := 0;
  NbHeureTravApMidi := 0;
  NbHDemiJournee := AFConversionUnite('J', 'H', 0.5);
  (* PL et CCO le 27/10/2004 : les règles de calendrier avaient été faites uniquement pour la gestion des 35h00
  et les notions de jour férié travaillé et de dimanche travaillé est gérable directement dans le calendrier
    if (TobCalendrierRegle <> nil) then
    begin
      // La Tob étant triée en ordre décroissant, on récupère la première date d'effet juste inférieure
      // à la date saisie et l'indice de la première TOB calendrier valable à la date d'effet
      vDtDerniereDateEffet := DetDerniereDateEffetCalendrier(pDtWork, i);
      if (i <> -1) then
      begin
        if (TobCalendrierRegle.Detail[i].GetValue('ACG_DATEEFFET') = vDtDerniereDateEffet) then
        begin
          if (pDtWork >= TobCalendrierRegle.Detail[i].GetValue('ACG_DATEEFFET')) then
          // Ce qui est normalement forcement le cas
          begin
            // Si c'est un jour férié et que ce n'est pas permis
            if (JourFerie.TestJourFerie(pDtWork) and (TobCalendrierRegle.Detail[i].GetValue('ACG_FERIEPERMIS')='-'))
            // Si c'est un dimanche et que ce n'est pas permis
               or ((DayOfWeek(pDtWork) = 1) and (TobCalendrierRegle.Detail[i].GetValue('ACG_DIMPERMIS')='-')) then
              Result:= false;
          end;
        end;
      end;
    end;
  *)

  if (TobCalendrier <> nil) and Result then
  begin
    // on recherche si le jour est travaillé
    for i := 0 to TobCalendrier.Detail.count - 1 do
    begin
      vDtCourante := TobCalendrier.Detail[i].GetValue('ACA_DATE');
      // cas 1
      // le jour est saisi dans le calendrier
      if (vDtCourante = pDtWork) then
      begin
        NbHeure := TobCalendrier.Detail[i].GetDouble('ACA_DUREE');
        NbHeureTravMatin := TobCalendrier.Detail[i].GetDouble('ACA_HEUREFIN1') - TobCalendrier.Detail[i].GetDouble('ACA_HEUREDEB1');
        NbHeureTravApMidi := TobCalendrier.Detail[i].GetDouble('ACA_HEUREFIN2') - TobCalendrier.Detail[i].GetDouble('ACA_HEUREDEB2');
        // si c'est un jour férié et qu'on ne peut pas travailler les jours fériés
        if gobj_JourFerie.TestJourFerie(pDtWork) and
          (TobCalendrier.Detail[i].GetValue('ACA_FERIETRAVAIL') = '-') then
        begin
          Result := false;
          break;
        end;

        // s'il y a une durée, alors le jour est travaillé
        if (TobCalendrier.Detail[i].GetValue('ACA_JOUR') = 0) then
        begin
          if (NbHeure > 0) then
          begin
            Result := true;
            break;
          end
          else
          begin
            Result := false;
            break;
          end;
        end
        else
        begin
          Result := false;
          break;
        end;
      end
        // cas 2
        // jours de la semaine dans le calendrier
      else
      begin
        vInJourSemaine := DayOfWeek(pDtWork) - 1;
        if (vInJourSemaine = 0) then
          vInJourSemaine := 7;
        if ((TobCalendrier.Detail[i].GetValue('ACA_JOUR') = vInJourSemaine) and
          (vDtCourante = iDate1900)) then
        begin
          if (TobCalendrier.Detail[i].GetValue('ACA_DUREE') > 0) then
            // S'il y a du temps de travail prévue sur ce jour de semaine
          begin
            if not gobj_JourFerie.TestJourFerie(pDtWork) then
            begin
              NbHeure := TobCalendrier.Detail[i].GetDouble('ACA_DUREE');
              NbHeureTravMatin := TobCalendrier.Detail[i].GetDouble('ACA_HEUREFIN1') - TobCalendrier.Detail[i].GetDouble('ACA_HEUREDEB1');
              NbHeureTravApMidi := TobCalendrier.Detail[i].GetDouble('ACA_HEUREFIN2') - TobCalendrier.Detail[i].GetDouble('ACA_HEUREDEB2');
              result := true;
            end
            else
            begin
              result := false;
            end;
            (*PL le  27/10/04 : le problème du jour férié se pose de 2 façons :
                - soit un horaire particulier est saisi ce jour férié et c'est traité dans le bloc d'avant (au jour le jour)
                - soit aucun horaire particulier n'est traité, et on dépend alors de l'horaire par semaine
                          if JourFerie.TestJourFerie(pDtWork) then
                          // Si c'est un jour férié
                          begin
                            // Le jour n'est pas travaillé et on sort de la boucle
                            result := false;
                            break;
                          end
                          else
                          begin*)
                            // Si ce n'est pas un jour férié
                            //result := true;
                            // break; PL le 04/10/04 : et on ne sort pas de la boucle parce qu'il faut continuer
                            // à scruter le calendrier au cas où un jour précis aurait été choisi comme non travaillé
            //              end;
          end
          else
          begin
            // S'il n'y a aucun temps de travail prévu sur ce jour de semaine
            result := false;
            // break; PL le 04/10/04 : et on ne sort pas de la boucle parce qu'il faut continuer
            // à scruter le calendrier au cas où un jour précis aurait été choisi comme travaillé
          end;
        end;
      end;
    end; // fin du for
  end;

  // si le jour est travaillé dans le calendrier
  // ou qu'il n'est travaillé qu'une demi journée
  // on verifie que ce n'est pas un jour de congés ou rtt dans la paie  (ou maladie...)
  if Result and (TobAbsenceSalarie <> nil) then
  begin
    TobAbsenceDuJour := AbsenceSalarieD1Date(pDtWork);
    if (TobAbsenceDuJour <> nil) then
    begin
      NbHAbsence := TobAbsenceDuJour.GetDouble('PCN_HEURES');
      NbJAbsence := TobAbsenceDuJour.GetDouble('PCN_JOURS'); //AB-200509-
      if bAxeHeures then
      begin
        // Plusieurs cas : 1- des heures ont été prises mais réparties sur la journée
        if (TobAbsenceDuJour.GetValue('PCN_DEBUTDJ') <> TobAbsenceDuJour.GetValue('PCN_FINDJ')) then
        begin
          //NbHeure := NbHeure - NbHAbsence
          //AB-200509 - NbHAbsence heures d'absences pour la période de NbJAbsence jours
          NbHeure := NbHeure - NbHAbsence / NbJAbsence;
          if (NbHeure < 0) then
          begin
            NbHeure := 0;
            Result := false;
          end
        end
        else
          // 2- on a pris que des heures le matin
          if (TobAbsenceDuJour.GetValue('PCN_DEBUTDJ') = 'MAT') and (TobAbsenceDuJour.GetValue('PCN_FINDJ') = 'MAT') then
        begin
          NbHeureTravMatin := NbHeureTravMatin - NbHAbsence;
          if (NbHeureTravMatin < 0) then
            NbHeureTravMatin := 0;

          NbHeure := NbHeureTravMatin + NbHeureTravApMidi;
        end
        else
          // 3- on a pris que des heures l'après midi
          if (TobAbsenceDuJour.GetValue('PCN_DEBUTDJ') = 'PAM') and (TobAbsenceDuJour.GetValue('PCN_FINDJ') = 'PAM') then
        begin
          NbHeureTravApMidi := NbHeureTravApMidi - NbHAbsence;
          if (NbHeureTravApMidi < 0) then
            NbHeureTravApMidi := 0;

          NbHeure := NbHeureTravMatin + NbHeureTravApMidi;
        end;
      end
      else
      begin
        // Plusieurs cas : 1- des heures ont été prises mais réparties sur la journée
        if (TobAbsenceDuJour.GetValue('PCN_DEBUTDJ') <> TobAbsenceDuJour.GetValue('PCN_FINDJ')) then
        begin
          NbHeure := 0;
          Result := false;
        end
        else
          // 2- on a pris que des heures le matin
          if (TobAbsenceDuJour.GetValue('PCN_DEBUTDJ') = 'MAT') and (TobAbsenceDuJour.GetValue('PCN_FINDJ') = 'MAT') then
        begin
          if (NbHeureTravApMidi > 0) then
            NbHeure := NbHDemiJournee
          else
          begin
            Result := false;
            NbHeure := 0;
          end;
        end
        else
          // 3- on a pris que des heures l'après midi
          if (TobAbsenceDuJour.GetValue('PCN_DEBUTDJ') = 'PAM') and (TobAbsenceDuJour.GetValue('PCN_FINDJ') = 'PAM') then
        begin
          if (NbHeureTravMatin > 0) then
            NbHeure := NbHDemiJournee
          else
          begin
            NbHeure := 0;
            Result := false;
          end;
        end;
      end;
    end;
  end;

  // Récapitulatif pour la gestion par jour :
  // PL le 15/11/04 : on choisit de gérer les absences à la demi journée de la façon suivante :
    // Si le jour est travaillé, qu'on est sur l'axe des absences et que le calcul de la demi journée n'est pas déjà fait dans les absences
      // alors si le nb d'heures travaillées est inférieur ou égal à 0.5j, alors on renvoie 0.5j  (et >0)
      // sinon on renvoie 1 j
  if (Result and not bAxeHeures and (NbHeure > 0)) then
  begin
    if (NbHeure <= NbHDemiJournee) then
    begin
      NbHeure := NbHDemiJournee;
    end
    else
      NbHeure := 2 * NbHDemiJournee;
  end;
end;

{$IFDEF AFFAIRE}
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 30/12/2002
Modifié le ... :
Description .. : Recherche du calendrier des ressources pour affecter les tâches
                 26/07/2005
                 modification de la méthode de décalage en prenant en compte
                 les éléments déjà dans la liste et la planification déjà effectuée

                 algorithme : on vérifie pour chaque élément de la liste
                 - si il n'est pas déjà dans la liste
                 - si il n'est pas deja travaille pour une autre affaire
                 - si c'est un jour travaillé

Mots clefs ... :
*****************************************************************}

procedure TAFO_Ressource.ControleListeEtDecale(var pListeJours: TListDate; pRdDuree: Double; pInNbJoursDecal: Integer; pStMethodeDecal: string; pTobJoursOccupes: Tob; pStAffaire, pStRes: string; pTache: RecordTache);
var
  i: Integer;
  vStJourListe: TDateTime;
  vStJourTrouve: TdateTime;
  vStJourEnCours: TdateTime;
  vBoTrouve: Boolean;

begin
  for i := pListeJours.Count - 1 downto 0 do
  begin
    vStJourListe := pListeJours.Items[i];
    vStJourEnCours := vStJourListe;
    vStJourTrouve := Idate1900;
    vBoTrouve := False;

    if PasDansListe(vStJourEnCours, pListeJours, False) and
      JourPasOccupe(vStJourEnCours, pTobJoursOccupes) and
      PeriodeTravaillee(vStJourEnCours, pRdDuree, vStJourTrouve) then
      vBoTrouve := True
    else
    begin
      vStJourTrouve := vStJourEnCours;
      case DecalageJour(pListeJours, pInNbJoursDecal, pStMethodeDecal, pRdDuree, vStJourTrouve, pTobJoursOccupes) of
        cInConserver:
          begin
            vBoTrouve := true;
            vStJourTrouve := vStJourEnCours;
          end;
        cInAnnuler:
          begin
            vBoTrouve := false;
          end;
        cInDecaler:
          begin
            if vStJourEnCours <> vStJourTrouve then
              vBoTrouve := true
            else
              vBoTrouve := false;
          end;
      end; // Case
    end;

    //cas 1 : conserver
    // rien

    // cas 2
    // supprime la planification pour cette ressource
    // et tenter le remplacement de ressource
    if not vBoTrouve then
    begin
      // C.B 29/07/2005
      // écriture dans le journal d''événements des jours annulés
      //AddJournal('ANN', PStAffaire, pTache.StTiers, pTache.StLibTache, pStRes, pListeJours.Strings[i], '');

			// 'Affaire : %s. Tiers : %s. Intervention %s non générée le %s pour la ressource %s.',
    	{1'Affaire : %s. Tiers : %s. Intervention %s non générée le %s pour la ressource %s.',}
//Modif FV AddJournal(fSlRapportAnn, format(TraduitGa(TexteMessage[1]), [PStAffaire, pTache.StTiers, pTache.StLibTache, pListeJours.Strings[i], pStRes]));
      pListeJours.Delete(i);

      // recherche de la ressource de remplacement
      // si paie serialisée
      {else if GetParamSoc ('SO_AFRESSREMP') then
      begin
      end

      // sinon effectation de la ressource de remplacement dans les paramsoc
      else
      begin
        //GetParamSoc ('SO_AFRESSREMP');
      end;
      }
    end

      // cas 3 decalage
      // un autre jour est trouvé
    else
      if vBoTrouve and (vStJourEnCours <> vStJourTrouve) and (vStJourTrouve <> IDate1900) then
    begin
      //AddJournal('DEC', PStAffaire, pTache.StTiers, pTache.StLibTache, pStRes, pListeJours.Strings[i], vStJourTrouve);
			//'Affaire : %s. Tiers : %s. Intervention %s décalée du %s au %s pour la ressource %s.',
    	{2'Affaire : %s. Tiers : %s. Intervention %s décalée du %s au %s pour la ressource %s.',}
//Modif FV AddJournal(fSlRapportDec, format(TraduitGa(TexteMessage[2]), [PStAffaire, pTache.StTiers, pTache.StLibTache, pListeJours.Strings[i], vStJourTrouve, pStRes]));

      pListeJours.Delete(i);
      pListeJours.add(vStJourTrouve);
    end;
  end; // end for
end;

{function TAFO_Ressource.GetRessourceRemplacement : TAFO_Ressource;
var
  Q : TQuery;
  S : String;

begin

  S := 'SELECT  ';
  Q := OpenSQL (S, true);

  Try
    if (Not Q.EOF) then
    begin

      // chargement de la ressource de remplacement

    end
    else
      result := nil;

  Finally
    Ferme(Q);
  End;
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 24/03/2003
Modifié le ... :   /  /
Description .. : On décale le jour de depart selon la regle choisie en parametre
                 retour : 0 conserve la valeur
                          1 annuler la planification
                          2 decale le jour
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.DecalageJour(var pListeJours: TListDate; pInNbJoursDecal: Integer; pStMethodeDecal: string; pRdDuree: Double; var pStJourEnCours: TdateTime; pTobJoursOccupes: Tob): Integer;
var
  i: integer;
  vBoTrouve: Boolean;
  vStJourEnCours: TdateTime;

begin
  i := 0;
  result := cInAnnuler;
  vBoTrouve := false;

  // regles de report
  if pStMethodeDecal = cStAnnule then
    result := cInAnnuler

  else
    if pStMethodeDecal = cStJour then
    result := cInConserver

  else
    if pStMethodeDecal = cStJourPrec then
  begin
    result := cInDecaler;
    vStJourEnCours := pStJourEnCours - 1;
    while (not vBoTrouve) and (i < pInNbJoursDecal - 1) do
    begin
      if PasDansListe(vStJourEnCours, pListeJours, True) and
        JourPasOccupe(vStJourEnCours, pTobJoursOccupes) and
        PeriodeTravaillee(vStJourEnCours, pRdDuree, pStJourEnCours) then
      begin
        vBoTrouve := true;
        break;
      end
      else
      begin
        i := i + 1;
        vStJourEnCours := vStJourEnCours - 1;
      end;
    end;
  end

  else
    if pStMethodeDecal = cStJourSuiv then
  begin
    result := cInDecaler;
    vStJourEnCours := pStJourEnCours + 1;
    while (not vBoTrouve) and (i < (pInNbJoursDecal - 1)) do
    begin
      if PasDansListe(vStJourEnCours, pListeJours, True) and
        JourPasOccupe(vStJourEnCours, pTobJoursOccupes) and
        PeriodeTravaillee(vStJourEnCours, pRdDuree, pStJourEnCours) then
      begin
        vBoTrouve := true;
        break;
      end
      else
      begin
        i := i + 1;
        vStJourEnCours := vStJourEnCours + 1;
      end;
    end;
  end;

  // si on devait decaler et qu'on n'a pas pu, on annule la planification
  if (not vBoTrouve) and (result = cInDecaler) then
    result := cInAnnuler;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : si l'ensemble des jours de la periode sont travaillés
                 la période est travaillée
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.PeriodeTravaillee(pStCurJour: TDateTime; pRdDuree: Double; var pStJourTrouve: TdateTime): Boolean;
var
  i: Integer;
  vDtJourEnCours: TDateTime;
  vInDuree: Integer;
  NbHeure: double;

begin

  ChargeCalendriers(True);
  vDtJourEnCours := pStCurJour;
  pStJourTrouve := IDate1900;
  result := true;
  if (frac(pRdDuree) <> 0) then
    vInDuree := trunc(pRdDuree) + 1
  else
    vInDuree := trunc(pRdDuree);

  for i := 0 to vInDuree - 1 do
  begin
    if JourTravaille(vDtJourEnCours, NbHeure, false) then
    begin
      // on retourne dans pStJourTrouve le premier jour de decalage trouvé
      if pStJourTrouve = IDate1900 then
        pStJourTrouve := vDtJourEnCours;
      vDtJourEnCours := vDtJourEnCours + 1;
    end
    else
    begin
      result := false;
      break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/07/2005
Modifié le ... :
Description .. : Chargement des journées planifiées pour une ressource sur
                 une période donnée pour ne pas planifier une aurte intervention
                 le même jour
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.LoadListeJoursOccupes(pStRes: string; pDtDateDeb, pDtDateFin: TDateTime): Tob;
var
  vSt: string;

begin
  vSt := 'SELECT APL_DATEDEBPLA, APL_DATEFINPLA FROM AFPLANNING ';
  vSt := vSt + ' WHERE APL_RESSOURCE = "' + pStRes + '"';
  vSt := vSt + ' AND APL_DATEDEBPLA >= "' + usDateTime(pDtDateDeb) + '"';
  vSt := vSt + ' AND APL_DATEDEBPLA <= "' + usDateTime(pDtDateFin) + '"';
  result := TOB.Create('AFPLANNING', nil, -1);
  result.LoadDetailFromSQL(vSt);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/07/2005
Modifié le ... :
Description .. : quand on passe dans cette fonction, la liste à été constituée
                 on teste donc si la date est 2 fois dans la stringlist
                 pBoAjout permet de savoir si on teste :
                 - une valeur à ajouter dans la liste, dans ce cas une seule
                   présence dans la liste indique deja dans la liste
                 - une valeur est en doublon dans la liste, et donc a decaler,
                   dans ce cas, la présence de 2 fois la date dans la liste
                   indique qu'il y a un doublon dans la liste
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.PasDansListe(vDtCurDate: TDateTime; pSLJours: TListDate; pBoAjout: Boolean): Boolean;
var
  pSLTest: TListDate;
  vInIndex: Integer;

begin
  result := True;
  if not GetParamsocSecur('SO_AFSURBOOKING', False) then
  begin
    pSLTest := TListDate.create;
    try
      pSLTest.Assign(pSLJours);
      vInIndex := pSLTest.indexof(vDtCurDate);

      // tester si on ajoute une valeur dans la liste
      if pBoAjout then
      begin
        if vInIndex <> -1 then
          result := False
        else
          result := True;
      end

        // tester si valeurs en doublons dans la liste
      else
      begin
        if vInIndex <> -1 then
        begin
          pSLTest.delete(vInIndex);
          vInIndex := pSLTest.indexof(vDtCurDate);
          if vInIndex <> -1 then
            result := False
          else
            result := True;
        end
        else
          result := True;
      end;
    finally
      pSLTest.Free;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/07/2005
Modifié le ... :
Description .. : teste si la date n'est pas déjà occupée par autre chose
Mots clefs ... :
*****************************************************************}

function TAFO_Ressource.JourPasOccupe(vDtCurDate: TDateTime; pTobJoursOccupes: Tob): Boolean;
var
  i: Integer;

begin
  result := True;
  if not GetParamsocSecur('SO_AFSURBOOKING', False) then
    for i := 0 to pTobJoursOccupes.detail.count - 1 do
      if (DateToStr(vDtCurDate) >= pTobJoursOccupes.detail[i].GetString('APL_DATEDEBPLA')) and
        (DateToStr(vDtCurDate) <= pTobJoursOccupes.detail[i].GetString('APL_DATEFINPLA')) then
      begin
        result := False;
        Break;
      end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/07/2005
Modifié le ... :
Description .. : tablette des événement GCTYPEEVENEMENT
Mots clefs ... :
*****************************************************************}
{procedure TAFO_Ressource.CreeJournaux;
var
  TobJnal: Tob;
  Indice: integer;
  QIndice: TQuery;

begin
  Indice := 0;

  QIndice := OpenSQL ('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', true);
  if not QIndice.Eof then
    Indice := QIndice.Fields[0].AsInteger + 1;
  Ferme (QIndice);

  // journal décalé
  if fMMRapportDec.Lines.count > 0 then
  begin
    TobJnal := Tob.Create ('JNALEVENT', nil, -1);
    try
      TOBJnal.SetString ('GEV_TYPEEVENT', 'PLG');
      TOBJnal.SetString ('GEV_LIBELLE', TexteMessage[3]);
      TOBJnal.SetDateTime ('GEV_DATEEVENT', Now);
      TOBJnal.SetString ('GEV_UTILISATEUR', V_PGI.User);
      TOBJnal.SetInteger ('GEV_NUMEVENT', Indice);
      TOBJnal.SetString ('GEV_ETATEVENT', 'OK');
      TOBJnal.SetString ('GEV_BLOCNOTE', fMMRapportDec.Lines.Text);
      TOBJnal.InsertDB (nil);
    finally
      TobJnal.free;
    end;
  end;

  // journal non généré
  if fMMRapportAnn.Lines.count > 0 then
  begin
    Indice := Indice + 1;
    TobJnal := Tob.Create ('JNALEVENT', nil, -1);
    try
      TOBJnal.SetString ('GEV_TYPEEVENT', 'PLG');
      TOBJnal.SetString ('GEV_LIBELLE', TexteMessage[4]);
      TOBJnal.SetDateTime ('GEV_DATEEVENT', Now);
      TOBJnal.SetString ('GEV_UTILISATEUR', V_PGI.User);
      TOBJnal.SetInteger ('GEV_NUMEVENT', Indice);
      TOBJnal.SetString ('GEV_ETATEVENT', 'OK');
      TOBJnal.SetString ('GEV_BLOCNOTE', fMMRapportAnn.Lines.Text);
      TOBJnal.InsertDB (nil);
    finally
      TobJnal.free;
    end;
  end;
end;
}

{procedure TAFO_Ressource.AddJournal(pStType, PStAffaire, pStTiers, pStInterv, pStRes, pStDtSource, pStDtDest : String);
begin
  if pStType = 'DEC' then
  begin
    fMMRapportDec.Lines.Append(format(TraduitGa(TexteMessage[2]), [PStAffaire, pStTiers, pStInterv, pStDtSource, pStDtDest, pStRes]));
    fMMRapportDec.Lines.Append('');
  end
  else if pStType = 'ANN' then
  begin
{procedure TAFO_Ressource.AddJournal(pStType, PStAffaire, pStTiers, pStInterv, pStRes, pStDtSource, pStDtDest : String);
    fMMRapportAnn.Lines.Append (format(TraduitGa(TexteMessage[1]), [PStAffaire, pStTiers, pStInterv, pStDtSource, pStRes]));
    fMMRapportAnn.Lines.Append ('');
  end;
end;                    
}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/10/2002
Modifié le ... :
Description .. : si on est en ajout, on ne cree les lignes que depuis la
                 derniere date de génération
Mots clefs ... :
*****************************************************************}

procedure TAFO_Ressource.ControleLastDateGene(var pListeJours: TListDate; pLastDateGene: TdateTime; pBoAjout: Boolean);
var
  i: Integer;
begin
  if pBoAjout then
    for i := pListeJours.count - 1 downto 0 do
    begin
      if pListeJours.Items[i] <= pLastDateGene then
        pListeJours.delete(i);
    end;
end;

{$ENDIF} // AFFAIRE

(******************************************************************************)
// ChargeCalendriers : Charge les calendriers liés à la ressource
//
//
// entrée : AcbForcer à true pour écraser les précédents
(******************************************************************************)

procedure TAFO_Ressource.ChargeCalendriers(AcbForcer: boolean);
var
  QQ: TQuery;
  sReq, sStandCalend: string;
  bCalSpecif: boolean;
begin
  QQ := nil;
  bCalSpecif := false;
  if (tob_Champs = nil) then
    exit;
  // PL le 04/10/05 : on va chercher un calendrier par défaut maintenant
  //if (tob_Champs.GetValue('ARS_STANDCALEN')='') then exit;

  // On peut créer les tob des temps d'absence et de dispo en attendant qu'elles soient remplies
  if (TobAbsencesParPeriode = nil) then
    TobAbsencesParPeriode := TOB.Create('Les Absences', nil, -1);
  if (TobDisposParPeriode = nil) then
    TobDisposParPeriode := TOB.Create('Les Dispos', nil, -1);

  //
  // Calendrier
  //
  if (tob_Champs.GetValue('ARS_CALENSPECIF') = 'X') then
    bCalSpecif := true;

  if (TobCalendrier = nil) or (AcbForcer = true) then
  begin
    TobCalendrier.Free;
    TobCalendrier := nil;
{$IFNDEF PLUGIN}
    TobCalendrier := Tob.Create('CALENDRIER', nil, -1);
    TobCalendrier.InitValeurs(FALSE);
{$ENDIF !PLUGIN}
    // S'il y a un calendrier spécifique...
    if bCalSpecif then
    begin
      try
        sStandCalend := tob_Champs.GetValue('ARS_STANDCALEN');
        if (sStandCalend <> '') then
        begin
          sReq := 'SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="' + sStandCalend
            + '" AND ACA_RESSOURCE="' + tob_Champs.GetValue('ARS_RESSOURCE') + '"';
          // SELECT * justifié car on a besoin de tous les champs pour les controles en aval
          // et le nombre d'enregistrements pour une ressource est restreint
{$IFDEF PLUGIN}
          TobCalendrier := PlugIn.OpenSelectInCache(sReq); //AB-200609-
{$ELSE !PLUGIN}
          QQ := OpenSQL(sReq, true,-1,'',true);
          if (not QQ.EOF) then
            TobCalendrier.LoadDetailDB('', '', '', QQ, false);
{$ENDIF !PLUGIN}
        end;

        if ((TobCalendrier<> nil) and (TobCalendrier.detail.count = 0)) or (sStandCalend = '') then
          bCalSpecif := false;

      finally
        Ferme(QQ);
        QQ := nil;
      end;
    end;

    // S'il n'y a pas de calendrier spécifique
    if (not bCalSpecif) then
    begin
      // Soit il y a un calendrier standard d'affecté à la ressource, et on prend celui là
      sStandCalend := tob_Champs.GetValue('ARS_STANDCALEN');
      if (sStandCalend <> '') then
      begin
        try
          //sReq := 'SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="' + tob_Champs.GetValue('ARS_STANDCALEN')
          //      + '" AND ACA_RESSOURCE="***"';
          // PL le 04/10/05 : on affine la requête pour tenir compte des spécificité de la paie
          sReq := 'SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="' + sStandCalend
            + '" AND ACA_RESSOURCE="***" AND ACA_SALARIE="***"';
          // SELECT * justifié car on a besoin de tous les champs pour les controles en aval
          // et le nombre d'enregistrements pour une ressource est restreint
{$IFDEF PLUGIN}
          TobCalendrier := PlugIn.OpenSelectInCache(sReq); //AB-200609-
{$ELSE !PLUGIN}
          QQ := OpenSQL(sReq, true,-1,'',true);
          if (not QQ.EOF) then
            TobCalendrier.LoadDetailDB('', '', '', QQ, false);
          //else // PL le 04/10/05 : on va chercher un calendrier par défaut maintenant
            //begin
              //TobCalendrier.Free;
              //TobCalendrier := nil;
            //end;
{$ENDIF !PLUGIN}
        finally
          Ferme(QQ);
          QQ := nil;
        end;
      end;

      if (TobCalendrier = nil) or ((TobCalendrier <> nil) and (TobCalendrier.Detail.count = 0)) then
        // Soit il n'y a pas de calendrier standard et on va chercher le calendrier par défaut
      begin
        sStandCalend := GetParamSocSecur('SO_AFSTANDCALEN', '');
        if (sStandCalend <> '') then
        begin
          try
            sReq := 'SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="' + sStandCalend
              + '" AND ACA_RESSOURCE="***" AND ACA_SALARIE="***"';
            // SELECT * justifié car on a besoin de tous les champs pour les controles en aval
            // et le nombre d'enregistrements pour une ressource est restreint
{$IFDEF PLUGIN}
            TobCalendrier := PlugIn.OpenSelectInCache(sReq); //AB-200609-
{$ELSE}
            QQ := OpenSQL(sReq, true,-1,'',true);
            if (not QQ.EOF) then
              TobCalendrier.LoadDetailDB('', '', '', QQ, false);
{$ENDIF !PLUGIN}
          finally
            Ferme(QQ);
            QQ := nil;
          end;
        end;
      end;

      if (TobCalendrier <> nil) and (TobCalendrier.Detail.count = 0) then
        // Si après tout ça on n'a toujours pas de calendrier, on libère la TOB
      begin
        TobCalendrier.Free;
        TobCalendrier := nil;
      end;

    end;
  end;

  //
  // Calendrier Regle
  //
  // Attention : la variable bCalSpecif a pu évoluer depuis l'init précédente
  if (tob_Champs.GetValue('ARS_CALENSPECIF') = 'X') then
    bCalSpecif := true;

  if (TobCalendrierRegle = nil) or (AcbForcer = true) then
  begin
    TobCalendrierRegle.Free;
    TobCalendrierRegle := nil;
{$IFNDEF PLUGIN}
    TobCalendrierRegle := Tob.Create('CALENDRIERREGLE', nil, -1);
    TobCalendrierRegle.InitValeurs(FALSE);
{$ENDIF !PLUGIN}
    if bCalSpecif then
    begin
      try
        // C'est important de laisser l'ordre décroissant par date d'effet, car on dispose ainsi du calendrier le
        // plus récent en premier
        sReq := 'SELECT * FROM CALENDRIERREGLE WHERE ACG_STANDCALEN="' + tob_Champs.GetValue('ARS_STANDCALEN')
          + '" AND ACG_RESSOURCE="' + tob_Champs.GetValue('ARS_RESSOURCE') + '" ORDER BY ACG_DATEEFFET DESC';
        // SELECT * justifié car on a besoin de tous les champs pour les controles en aval
        // et le nombre d'enregistrements pour une ressource est restreint
{$IFDEF PLUGIN}
        TobCalendrierRegle := PlugIn.OpenSelectInCache(sReq); //AB-200609-
        if (TobCalendrierRegle = nil) or ((TobCalendrierRegle <> nil) and (TobCalendrierRegle.detail.count=0)) then
          bCalSpecif := false;
{$ELSE}
        QQ := OpenSQL(sReq, true,-1,'',true);
        if (not QQ.EOF) then
          TobCalendrierRegle.LoadDetailDB('', '', '', QQ, false)
        else
          bCalSpecif := false;
{$ENDIF !PLUGIN}
      finally
        Ferme(QQ);
        QQ := nil;
      end;
    end;

    if not bCalSpecif then
    begin
      try
        // C'est important de laisser l'ordre décroissant par date d'effet, car on dispose ainsi du calendrier le
        // plus récent en premier
        sReq := 'SELECT * FROM CALENDRIERREGLE WHERE ACG_STANDCALEN="' + tob_Champs.GetValue('ARS_STANDCALEN')
          + '" AND ACG_RESSOURCE="***" ORDER BY ACG_DATEEFFET DESC';
        // SELECT * justifié car on a besoin de tous les champs pour les controles en aval
        // et le nombre d'enregistrements pour une ressource est restreint
{$IFDEF PLUGIN}
        TobCalendrierRegle := PlugIn.OpenSelectInCache(sReq); //AB-200609-
{$ELSE}
        QQ := OpenSQL(sReq, true,-1,'',true);
        if (not QQ.EOF) then
          TobCalendrierRegle.LoadDetailDB('', '', '', QQ, false)
        else
        begin
          TobCalendrierRegle.Free;
          TobCalendrierRegle := nil;
        end;
{$ENDIF !PLUGIN}
      finally
        Ferme(QQ);
      end;
    end;
  end;

  //
  // Absence des salariés
  //
  if (AcbForcer = true) then
    giAbsencePaie := 0; // On force le rechargement des absences salarié
  if PaieUtilisee then
  begin
    ChargeAbsencesSalarie(AcbForcer);
  end;
end;
 
{function TAFO_Ressource.AUnCalendrier : Boolean;
begin
  if (TobCalendrier = nil) then
  begin
    ChargeCalendriers(true);
    if (TobCalendrier = nil) then
      result := False
    else
      result := True;
  end
  else
    result := True;
end;}

(******************************************************************************)
// ChargeAbsencesSalarie : Charge les absences de la ressource depuis la table de la paie
//
//
// entrée : AcbForcer à true pour écraser les précédents
(******************************************************************************)
procedure TAFO_Ressource.ChargeAbsencesSalarie(AcbForcer: boolean);
var
  sReq: string;
  QQ: TQuery;
begin

  if (TobAbsenceSalarie <> nil) and not AcbForcer then
    exit;

  TobAbsenceSalarie.Free;
  TobAbsenceSalarie := nil;

  try
    sReq := 'SELECT PCN_SALARIE,PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS,' +
      'PCN_TYPEMVT,PCN_ORDRE,PCN_JOURS,PCN_HEURES,PCN_DEBUTDJ,PCN_FINDJ,' +
      'PMA_LIBELLE,PMA_ABREGE,PMA_EDITION,PMA_PRISTOTAL FROM ABSENCESALARIE ' +

      //C.B 13/04/2006
      // ajout suite demande de la paie
      'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE=PCN_TYPECONGE ' +
      'WHERE PCN_SALARIE="' + tob_Champs.GetValue('ARS_SALARIE') +
	    //C.B 13/04/2006
  	  // ne pas ternir compte des absences annulées
    	'" AND PCN_ETATPOSTPAIE <> "NAN" ' +       // PL le 19/04/06 : oublie guillemets, ce commentaire est à virer après vérif

      //                  '" AND ((PCN_DATEDEBUTABS>="'+USDateTime(DateDeb-1)+'" '+
    //                  'AND PCN_DATEDEBUTABS<="'+USDateTime(DateFin)+'") '+ //PT29REP DateDeb-1 : Cas abs fin de mois anterieur, et 1er jour mois en cours repos salarié
    //                  'OR (PCN_DATEFINABS>="'+USDateTime(DateDeb-1)+'" '+
    //                  'AND PCN_DATEFINABS<="'+USDateTime(DateFin)+'" )) '+
		// PL le 19/04/06 : guillemets en trop, ce commentaire est à virer après vérif
    ' AND ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-") OR (PCN_TYPECONGE="PRI" AND PCN_TYPEMVT="CPA" AND PCN_MVTDUPLIQUE="-" )) ' + //PT35REP Ajout AND PCN_SENSABS="-"
    //            'AND PMA_EDITION="X" '+ // PL le 08/11/04 : vu avec la paie ce jour : ne sert à rien, juste pour voir 'absence dans les éedin
    'ORDER BY PCN_DATEDEBUTABS,PCN_DATEFINABS';
{$IFDEF PLUGIN}
    TobAbsenceSalarie := PlugIn.OpenSelectInCache(sReq); //AB-200609-
{$ELSE}
    TobAbsenceSalarie := Tob.Create('Les_Absences_Salarie', nil, -1);
    TobAbsenceSalarie.InitValeurs(FALSE);
    QQ := OpenSQL(sReq, true,-1,'',true);
    if (not QQ.EOF) then
      TobAbsenceSalarie.LoadDetailDB('', '', '', QQ, false);
{$ENDIF !PLUGIN}
  finally
    Ferme(QQ);
    QQ := nil;
  end;
end;

(******************************************************************************)
// AbsenceSalarieD1Date : cherche si une date est dans un intervalle d'absence de la
// ressource courante dans la table ABSENCESALARIE de la paie
//
//
// entrée : AcdDateTestee : date à tester
// sortie : renvoie la tob de la ligne d'absence si elle est trouvée, sinon nil
(******************************************************************************)

function TAFO_Ressource.AbsenceSalarieD1Date(AcdDateTestee: TDateTime): TOB;
var
  TOBDet: TOB;
  i: integer;
begin
  Result := nil;
  if PaieUtilisee then
  begin
    for i := 0 to (TobAbsenceSalarie.Detail.count - 1) do
      // Sur toutes les absences du salarie saisies dans la paie
    begin
      TOBDet := TobAbsenceSalarie.Detail[i];
      if (AcdDateTestee >= TOBDet.GetValue('PCN_DATEDEBUTABS')) and (AcdDateTestee <= TOBDet.GetValue('PCN_DATEFINABS')) then
      begin
        Result := TOBDet;
        break;
      end;
    end;
  end;
end;

destructor TAFO_Ressource.Destroy;
begin
  gsUniteValo := '';
  tob_Champs.Free;      
  Article.Free;
  TobRessourcePR.Free;
  TobCalendrier.Free;
  TobCalendrierRegle.Free;
  TobAbsenceSalarie.Free;
  TobAbsencesParPeriode.Free;
  TobDisposParPeriode.Free;
//{$IFDEF AFFAIRE}
  gobj_JourFerie.Free;
//{$ENDIF}
  inherited;
end;

(******************************************************************************)
// TempsAnnuelRessource : donne le temps de travail effectif de la ressource
//                        dans l'unite de saisie d'activite
//
// entrée : Type d'activite : REA ou PLA, Date de début et Date de fin d'analyse
(******************************************************************************)

function TAFO_Ressource.TempsAnnuelRessource(AcsTypeActivite: string; AcdDateDebut, AcdDateFin: TDateTime): double;
var
  sSelectR: string;
  QSelectR: TQuery;
  TobQte: TOB;
  i: integer;
  dQteCourante: double;
  sUniteCourante: string;

begin
  Result := 0;
  QSelectR := nil;
  if (tob_Champs = nil) then
    exit;

  sSelectR := 'SELECT SUM(ACT_QTE) AS ACT_QTE, ACT_UNITE FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + AcsTypeActivite
    + '" AND ACT_RESSOURCE="' + tob_Champs.GetValue('ARS_RESSOURCE')
    + '" AND ACT_DATEACTIVITE>="' + UsDateTime(AcdDateDebut)
    + '" AND ACT_DATEACTIVITE<="' + UsDateTime(AcdDateFin)
    + '" AND ACT_TYPEARTICLE="PRE" AND ACT_ACTIVITEEFFECT="X" GROUP BY ACT_UNITE';

  TobQte := TOB.Create('', nil, -1);
  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);
    if not QSelectR.EOF then
      // Si la sélection n'est pas vide
    begin

      TobQte.LoadDetailDB('Qte par an', '', '', QSelectR, False);

      for i := 0 to TobQte.Detail.count - 1 do
      begin
        sUniteCourante := TobQte.Detail[i].GetValue('ACT_UNITE');
        dQteCourante := TobQte.Detail[i].GetValue('ACT_QTE');
        Result := Result + AFConversionUnite(sUniteCourante, gsUniteActivite, dQteCourante);
      end;
    end;
  finally
    Ferme(QSelectR);
    TobQte.Free;
  end;
end;

(******************************************************************************)
// FonctionDeLaRessource
// Entrée : Date à laquelle on cherche la fonction
//
// Si la date d'entrée est vide, la fonction renvoi la liste des fonctions rangées
// dans l'ordre chronologique, l'indice 0 étant la plus récente et l'indice 3 la plus
// ancienne
// Si la date d'entrée est saisie, la fonction renvoi une liste avec pour seul élément
// la fonction à la date d'entrée (indice 0)
(******************************************************************************)

function TAFO_Ressource.FonctionDeLaRessource(AcdDateDeFonction: TDateTime): TStringList;
var
  vslFonctions, sDates: TStringList;
  i, j, k: integer;
begin
  Result := nil;

  vslFonctions := TStringList.Create;
  sDates := TStringList.Create;

  try
    sDates.Add(tob_Champs.GetValue('ARS_DATEFONC1'));
    vslFonctions.Add(tob_Champs.GetValue('ARS_FONCTION1'));
    sDates.Add(tob_Champs.GetValue('ARS_DATEFONC2'));
    vslFonctions.Add(tob_Champs.GetValue('ARS_FONCTION2'));
    sDates.Add(tob_Champs.GetValue('ARS_DATEFONC3'));
    vslFonctions.Add(tob_Champs.GetValue('ARS_FONCTION3'));
    sDates.Add(tob_Champs.GetValue('ARS_DATEFONC4'));
    vslFonctions.Add(tob_Champs.GetValue('ARS_FONCTION4'));

    // on trie la liste dans l'ordre des dates
    for k := 0 to 3 do
      for i := 3 downto 0 do
        for j := 3 downto 0 do
        begin
          if isvaliddate(sDates[i]) and isvaliddate(sDates[j]) then
            if (i < j) and (strtodate(sDates[i]) < strtodate(sDates[j])) then
            begin
              vslFonctions.exchange(i, j);
              sDates.exchange(i, j);
              break;
            end;
        end;

    if (AcdDateDeFonction = 0) or (AcdDateDeFonction = iDate1900) then
      // Si aucune date n'est fournie en entrée, on renvoie la liste des 4 dernières
      // fonctions pratiquées triées de la plus rescente=0 à la plus ancienne=3
    begin
      Result := vslFonctions;
      exit;
    end;

    // Une date d'entrée a été fournie, on cherche donc la fonction à la date juste inférieure
    // à cette date entrée
    try
      for i := 0 to 3 do
      begin
        if (AcdDateDeFonction >= strtodate(sDates[i])) then
        begin
          Result := TStringList.Create;
          TStringList(Result).Add(vslFonctions[i]);
          break;
        end;
      end;

    finally
      vslFonctions.Free;
    end;

  finally
    sDates.free;
  end;
end;

(******************************************************************************)
// PrixCourant
// Entrée : Date à laquelle on cherche la prix
//          Type de prix recherché (PR ou PV)
//          pour le PV : HT ou TTC
//
// Si la date d'entrée est vide, la fonction renvoi la liste des prix rangés
// dans l'ordre chronologique, l'indice 0 étant le plus récent et l'indice 3 le plus
// ancien
// Si la date d'entrée est saisie, la fonction renvoi une liste avec pour seul élément
// le prix à la date d'entrée (indice 0)
(******************************************************************************)

function TAFO_Ressource.PrixCourant(AcdDateDePrix: TDateTime; AcsTypePrix: string; AcbPVHT: boolean = true): TStringList;
var
  vslPrix, sDates: TStringList;
  i, j, k: integer;
begin
  Result := nil;

  vslPrix := TStringList.Create;
  sDates := TStringList.Create;

  try
    sDates.Add(tob_Champs.GetValue('ARS_DATEPRIX'));
    sDates.Add(tob_Champs.GetValue('ARS_DATEPRIX2'));
    sDates.Add(tob_Champs.GetValue('ARS_DATEPRIX3'));
    sDates.Add(tob_Champs.GetValue('ARS_DATEPRIX4'));
    if (AcsTypePrix = 'R') then
      // Prix de revient
    begin
      vslPrix.Add(tob_Champs.GetValue('ARS_TAUXREVIENTUN'));
      vslPrix.Add(tob_Champs.GetValue('ARS_TAUXREVIENTUN2'));
      vslPrix.Add(tob_Champs.GetValue('ARS_TAUXREVIENTUN3'));
      vslPrix.Add(tob_Champs.GetValue('ARS_TAUXREVIENTUN4'));
    end
    else
      // Prix de vente
    begin
      if AcbPVHT then
        // HT
      begin
        vslPrix.Add(tob_Champs.GetValue('ARS_PVHT'));
        vslPrix.Add(tob_Champs.GetValue('ARS_PVHT2'));
        vslPrix.Add(tob_Champs.GetValue('ARS_PVHT3'));
        vslPrix.Add(tob_Champs.GetValue('ARS_PVHT4'));
      end
      else
        // TTC
      begin
        vslPrix.Add(tob_Champs.GetValue('ARS_PVTTC'));
        vslPrix.Add(tob_Champs.GetValue('ARS_PVTTC2'));
        vslPrix.Add(tob_Champs.GetValue('ARS_PVTTC3'));
        vslPrix.Add(tob_Champs.GetValue('ARS_PVTTC4'));
      end;
    end;

    // on trie la liste dans l'ordre des dates
    for k := 0 to 3 do
      for i := 3 downto 0 do
        for j := 3 downto 0 do
        begin
          if isvaliddate(sDates[i]) and isvaliddate(sDates[j]) then
            if (i < j) and (strtodate(sDates[i]) < strtodate(sDates[j])) then
            begin
              vslPrix.exchange(i, j);
              sDates.exchange(i, j);
              break;
            end;
        end;

    if (AcdDateDePrix = 0) or (AcdDateDePrix = iDate1900) then
      // Si aucune date n'est fournie en entrée, on renvoie la liste des 4 derniers
      // prix triés du plus rescent=0 au plus ancien=3
    begin
      Result := vslPrix;
      exit;
    end;

    // Une date d'entrée a été fournie, on cherche donc le prix à la date juste inférieure
    // à cette date entrée
    try
      for i := 0 to 3 do
      begin
        if (AcdDateDePrix >= strtodate(sDates[i])) then
        begin
          Result := TStringList.Create;
          TStringList(Result).Add(vslPrix[i]);
          break;
        end;
      end;

    finally
      vslPrix.Free;
    end;

  finally
    sDates.free;
  end;
end;

(******************************************************************************)
// AssocieArticle
// Entrée : Code Article à associer
//
// Associe la TOB du code article en entree à la ressource courante
// Si le code article est vide, associe l'article par defaut de la ressource
// ATTENTION : supprime l'ancien article associe
(******************************************************************************)
{$IFDEF AFFAIRE} //mcd 04/06/03 mis en ifdef affaire .. appeler depuis activité. intuile en paie

function TAFO_Ressource.AssocieArticle(AcsCodeArticle: string): boolean;
var
  tobArticle: TOB;
begin
  Result := false;
  // on reinitialise l'unite de temps de valorisation avec l'unite de la ressource
  gsUniteValo := tob_Champs.GetValue('ARS_UNITETEMPS');

  tobArticle := Tob.Create('ARTICLE', nil, -1);
  tobArticle.InitValeurs(FALSE);
  if (AcsCodeArticle = '') then
    // s'il n'y a pas d'article associé, on prend l'article de la ressource par defaut
    tobArticle.SelectDB('"' + tob_Champs.GetValue('ARS_ARTICLE') + '"', nil);
//else
//Modif FV    TrouverArticleSQL_GI(false, AcsCodeArticle, tobArticle, '', GetParamSocSecur('SO_AFFRAISCOMPTA', False)); // PL le 10/03/04 : on ne ramène les compléments que pour les frais en compta

  // Si on n'a toujours pas d'article associé, on ne peut rien faire
  if (tobArticle = nil) then
    exit;

  if (tobArticle.GetValue('GA_ARTICLE') = '') then
  begin
    tobArticle.Free;
    exit;
  end;

  // Sinon, on associe le nouvel article
  if (Article <> nil) then
    Article.Free;
  Article := tobArticle;
  Result := true;
end;
{$ENDIF}

(******************************************************************************)
// AssocieRessourcePR
// Entrée : Code Ressource
// AcbForcer : forcer la mise à jour
//
// Associe la TOB définie par les valeurs en entree à la ressource courante
// la TOB contient 1 ou plusieurs occurrences suivant les fonctions et le rang
(******************************************************************************)

function TAFO_Ressource.AssocieRessourcePR(AcsCodeRessource: string; AcbForcer: boolean): boolean;
var
  QSelectR: TQuery;
  sSelectR: string;
begin
  if (TobRessourcePR <> nil) and (AcbForcer = False) then
  begin
    Result := true;
    exit;
  end;

  if (TobRessourcePR = nil) then
    TobRessourcePR := TOB.Create('Les RessourcePRs', nil, -1)
  else
  begin
    TobRessourcePR.ClearDetail;
    TobRessourcePR.InitValeurs(False);
  end;

  // on selectionne tous les PR ou PV (suivant le type valo) par fonctions/rang correspondant à la ressource courante
  sSelectR := 'SELECT ARP_TYPEVALO, ARP_RESSOURCE, ARP_FONCTION, ARP_RANG, ARP_FAMILLENIV1, ARP_FAMILLENIV2, ARP_FAMILLENIV3, ARP_ARTICLE, ARP_FORMULE , ARP_DEPARTEMENT FROM RESSOURCEPR WHERE ARP_RESSOURCE="'
    + AcsCodeRessource + '"';

  QSelectR := nil;
  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);
    if not QSelectR.EOF then
      TobRessourcePR.LoadDetailDB('', '', '', QSelectR, false);

    Result := true;
  finally
    Ferme(QSelectR);
  end;
end;

(******************************************************************************)
//
// PRRessource : renvoie le prix de revient de la ressource
//
// Entrée : Date du prix cherché
//          TOB Article pour lequel on fait cette recherche
//          ou le code Article
//          Code de l'affaire liée
//
// Si une ligne de PR dans la table RESSOURCEPR correspond à l'article en entrée
// suivant l'ordre de priorité suivant
//                                     - identifiant article
//                                     - Familles d'article 1,2 et 3
//                                     - Familles d'article 1 et 2
//                                     - Familles d'article 1 et 3
//                                     - Familles d'article 2 et 3
//                                     - Familles d'article 1
//                                     - Familles d'article 2
//                                     - Familles d'article 3
//                                     - Département
// on applique la formule de calcul du PR associée à la ligne courante
// sinon,
// on lit le taux de revient unitaire de la table RESSOURCE
//
(******************************************************************************)

function TAFO_Ressource.CalculPrixRessource(AcdDate: TDateTime; AcsTypePrix: string; ActobArticle: TOB; ActobAffaires: TOB; AcsCodeArticle: string; AcsCodeAffaire: string = ''): double;
var
  sArticle: string;
  sFam123, sFam12, sFam13, sFam23, sFam1, sFam2, sFam3: string;
  sDep, sFormuleDep: string;
  {$IFDEF VER150}
  sFormule: hString;
  {$ELSE}
  sFormule: string;
  {$ENDIF}
  vResultat, v2: variant;
  dCoeffConvert: double;
{$IFDEF AFFAIRE}
  TOBAffaire: TOB;
{$ENDIF}
	TT: TOB;
  ListePrix: TStringList;
begin
  Result := 0;
  sArticle := '';
  sFam123 := '';
  sFam12 := '';
  sFam13 := '';
  sFam23 := '';
  sFam1 := '';
  sFam2 := '';
  sFam3 := '';
  sDep := '';
  sFormuleDep := '';
  sFormule := '';
  gDatePrix := AcdDate;

  if (ActobArticle = nil) then
  begin
{$IFDEF AFFAIRE} //mcd 04/03/03
    AssocieArticle(AcsCodeArticle);
{$ENDIF}
  end
  else
  begin
    if (Article <> nil) then
      Article.Free;

    Article := Tob.Create('ARTICLE', nil, -1);
    Article.Dupliquer(ActobArticle, true, true, false);
  end;

{$IFDEF AFFAIRE}
  if (AcsCodeAffaire <> '') then
  begin
    TobAffaire := Tob.Create('AFFAIRE', nil, -1);
    try
      if AFRemplirTOBAffaire(AcsCodeAffaire, TobAffaire, ActobAffaires) then
        sDep := TobAffaire.GetValue('AFF_DEPARTEMENT');

    finally
      TobAffaire.Free;
    end;
  end;
{$ENDIF}

  if (Article <> nil) or (sDep <> '') then
  begin
    if AssocieRessourcePR(tob_Champs.GetValue('ARS_RESSOURCE'), false) then
    begin
      // on cherche si une ligne de PR correspond à l'article en entrée, pour la ressource courante
      // Si aucune ligne de correspond, on cherche si les trois famille de l'article en entrée correspondent
      // aux trois familles d'une ligne PR ou
      // que les familles correspondent deux par deux ou une à une
      // Si rien ne correspond, on cherche si le département en entrée correspond au département d'une ligne PR
      TT := TobRessourcePR.FindFirst(['ARP_TYPEVALO', 'ARP_RESSOURCE'], [AcsTypePrix, tob_Champs.GetValue('ARS_RESSOURCE')], TRUE);
      while (TT <> nil) and (sArticle = '') do
      begin
        if (Article <> nil) then
        begin
          if (TT.GetValue('ARP_ARTICLE') = Article.GetValue('GA_ARTICLE')) then
            sArticle := TT.GetValue('ARP_FORMULE');

          if ((TT.GetValue('ARP_FAMILLENIV1') <> '')
            and (TT.GetValue('ARP_FAMILLENIV2') <> '')
            and (TT.GetValue('ARP_FAMILLENIV3') <> '')) then
            if ((TT.GetValue('ARP_FAMILLENIV1') = Article.GetValue('GA_FAMILLENIV1'))
              and (TT.GetValue('ARP_FAMILLENIV2') = Article.GetValue('GA_FAMILLENIV2'))
              and (TT.GetValue('ARP_FAMILLENIV3') = Article.GetValue('GA_FAMILLENIV3'))) then
              sFam123 := TT.GetValue('ARP_FORMULE');

          if ((TT.GetValue('ARP_FAMILLENIV1') <> '') and (TT.GetValue('ARP_FAMILLENIV2') <> '')) then
            if (TT.GetValue('ARP_FAMILLENIV1') = Article.GetValue('GA_FAMILLENIV1'))
              and (TT.GetValue('ARP_FAMILLENIV2') = Article.GetValue('GA_FAMILLENIV2')) then
              sFam12 := TT.GetValue('ARP_FORMULE');

          if ((TT.GetValue('ARP_FAMILLENIV1') <> '') and (TT.GetValue('ARP_FAMILLENIV3') <> '')) then
            if (TT.GetValue('ARP_FAMILLENIV1') = Article.GetValue('GA_FAMILLENIV1'))
              and (TT.GetValue('ARP_FAMILLENIV3') = Article.GetValue('GA_FAMILLENIV3')) then
              sFam13 := TT.GetValue('ARP_FORMULE');

          if ((TT.GetValue('ARP_FAMILLENIV2') <> '') and (TT.GetValue('ARP_FAMILLENIV3') <> '')) then
            if (TT.GetValue('ARP_FAMILLENIV2') = Article.GetValue('GA_FAMILLENIV2'))
              and (TT.GetValue('ARP_FAMILLENIV3') = Article.GetValue('GA_FAMILLENIV3')) then
              sFam23 := TT.GetValue('ARP_FORMULE');

          if (TT.GetValue('ARP_FAMILLENIV1') <> '') then
            if (TT.GetValue('ARP_FAMILLENIV1') = Article.GetValue('GA_FAMILLENIV1')) then
              sFam1 := TT.GetValue('ARP_FORMULE');

          if (TT.GetValue('ARP_FAMILLENIV2') <> '') then
            if (TT.GetValue('ARP_FAMILLENIV2') = Article.GetValue('GA_FAMILLENIV2')) then
              sFam2 := TT.GetValue('ARP_FORMULE');

          if (TT.GetValue('ARP_FAMILLENIV3') <> '') then
            if (TT.GetValue('ARP_FAMILLENIV3') = Article.GetValue('GA_FAMILLENIV3')) then
              sFam3 := TT.GetValue('ARP_FORMULE');

        end;

        if (sDep <> '') then
        begin
          if (TT.GetValue('ARP_DEPARTEMENT') = sDep) then
            sFormuleDep := TT.GetValue('ARP_FORMULE');
        end;

        TT := TobRessourcePR.FindNext(['ARP_TYPEVALO', 'ARP_RESSOURCE'],
          [AcsTypePrix, tob_Champs.GetValue('ARS_RESSOURCE')],
          TRUE);
      end;

      if (sArticle <> '') then
        // on a trouve une ligne de PR correspondant à l'article en entrée
        // la formule à appliquer est stockée dans la variable sArticle
        sFormule := sArticle
      else
        if (sFam123 <> '') then
        sFormule := sFam123
      else
        if (sFam12 <> '') then
        sFormule := sFam12
      else
        if (sFam13 <> '') then
        sFormule := sFam13
      else
        if (sFam23 <> '') then
        sFormule := sFam23
      else
        if (sFam1 <> '') then
        sFormule := sFam1
      else
        if (sFam2 <> '') then
        sFormule := sFam2
      else
        if (sFam3 <> '') then
        sFormule := sFam3
      else
        if (sDep <> '') then
        sFormule := sFormuleDep;

    end;
  end;

  if (sFormule <> '') then
    // on a trouvé un article ou une combinaison de famille d'articles correspondant
    // à l'article en entrée ou un département correspondant correspondant au département en entrée
    // Il faut appliquer cette formule pour obtenir le PR
  begin
    V_PGI.ErreurFormule := true;
    sFormule := '{' + sFormule + '}';
    vResultat := 0;
    if TestFormule(sFormule) then
    begin
      sFormule := '{"#.###,0"' + sFormule + '}'; //mcd 09/12/02
      vResultat := GFormule(sFormule, DonneesFormuleP, nil, 1);
      // PL le 14/01/03 : gestion des erreurs
      if (vResultat <> 'Erreur') and (vResultat <> '') then
      begin
        try
          v2 := VarAsType(vResultat, varDouble);
          Result := v2;
        except
          Result := 0;
        end;
      end;
    end;
  end
  else
    // on n'a pas trouvé d'article ni de combinaison de famille correspondant
    // à l'article en entrée ni de département, on prend le champ TAUXREVIENTUN de la ressource
    // ou PVHT ou TTC suivant le type de valo
  begin
    dCoeffConvert := 1;
    ListePrix := PrixCourant(AcdDate, AcsTypePrix, gbPVHT);
    if (gsUniteValo <> tob_Champs.GetValue('ARS_UNITETEMPS')) then
      // Si l'unite de valorisation est différente de l'unite de la ressource
      // on doit convertir en sortie
    begin
      dCoeffConvert := AFConversionUnite(tob_Champs.GetValue('ARS_UNITETEMPS'), gsUniteValo, 1);
      if dCoeffConvert = 0 then
        dCoeffConvert := 1;
    end;

    if (ListePrix <> nil) then
    begin
      if (ListePrix.Count <> 0) then
        Result := Valeur(ListePrix[0]) / dCoeffConvert;
      ListePrix.Free;
    end;
  end;

end;

(******************************************************************************)
// ExisteUnArticleDeValo : renvoie true si il existe une valorisation personnalisée
//                         de la ressource pour l'article passé en paramètre
//
// Entrée :     AcsTypeValo type de valorisation
//              AcsArticle article à tester
//
(******************************************************************************)

function TAFO_Ressource.ExisteUnArticleDeValo(AcsTypeValo, AcsArticle: string): boolean;
var
  T: TOB;
begin
  Result := false;
  if (AcsArticle = '') or (AcsTypeValo = '') then
    exit;

  if AssocieRessourcePR(tob_Champs.GetValue('ARS_RESSOURCE'), false) then
  begin
    T := TobRessourcePR.FindFirst(['ARP_TYPEVALO', 'ARP_RESSOURCE', 'ARP_ARTICLE'], [AcsTypeValo, tob_Champs.GetValue('ARS_RESSOURCE'), AcsArticle], TRUE);
    if (T <> nil) then
      Result := true;
  end;
end;

(******************************************************************************)
// DonneesFormuleP : renvoie un variant contenant la valeur du champ de la table
//                      RESSOURCE ou ARTICLE suivant le ChampFormule passé en entrée
//
// Entrée : ChampFomule contenant le code défini dans la formule correspondant
//          à un champ soit dans la table RESSOURCE, soit dans la table ARTICLE
(******************************************************************************)

function TAFO_Ressource.DonneesFormuleP(AcsChampFormule: Hstring): variant;
var
  sTypePrix: string;
  dTauxUnit, dTauxUnitPres, dCoeffConvert, dCoeffConvert2: double;
  ListePrix: TStringList;
begin
  ListePrix := nil;
  dTauxUnit := 0;
  dTauxUnitPres := 0;
  dCoeffConvert := 1; //dCoeffConvert2 := 1;
  // On détermine en quelle unité on traduit : soit dans l'unité de la ressource, soit dans l'unité de l'article associé
  if (Article <> nil) then
  begin
    if (AcsChampFormule = 'TAUXREVIENT') or (AcsChampFormule = 'TAUXVENTE') then
    begin
      sTypePrix := 'R';
      if (AcsChampFormule = 'TAUXVENTE') then
        sTypePrix := 'V';
      ListePrix := PrixCourant(gDatePrix, sTypePrix, gbPVHT);
    end;

    if (gsUniteValo = tob_Champs.GetValue('ARS_UNITETEMPS')) then
      // Unité de la ressource
    begin
      dTauxUnit := tob_Champs.GetValue('ARS_TAUXUNIT');

      if (gsUniteValo = Article.GetValue('GA_QUALIFUNITEVTE')) then
        // Si l'article est exprimé dans la meme unité que la ressource
        dTauxUnitPres := Article.GetValue('GA_PMRP')
      else
        // sinon, on traduit de l'unite de l'article vers l'unite de valorisation (ressource)
      begin
        dCoeffConvert2 := AFConversionUnite(Article.GetValue('GA_QUALIFUNITEVTE'), gsUniteValo, 1);
        if dCoeffConvert2 = 0 then
          dCoeffConvert2 := 1;
        dTauxUnitPres := Article.GetValue('GA_PMRP') / dCoeffConvert2;
      end;
    end
    else
      if (gsUniteValo = Article.GetValue('GA_QUALIFUNITEVTE')) then
      // Unité de l'article
    begin
      // l'unite de la ressource est forcement differente de l'unite de valorisation (article), sinon on serait passé dans le
      // premier test
      // on convertit donc la ressource dans l'unite de valorisation (article)
      dCoeffConvert := AFConversionUnite(tob_Champs.GetValue('ARS_UNITETEMPS'), gsUniteValo, 1);
      if dCoeffConvert = 0 then
        dCoeffConvert := 1;
      dTauxUnit := tob_Champs.GetValue('ARS_TAUXUNIT') / dCoeffConvert;

      dTauxUnitPres := Article.GetValue('GA_PMRP');
    end
    else
      // autre unité (unité d'activité par exemple)
    begin
      // Tout est converti dans l'unité de valorisation
      dCoeffConvert := AFConversionUnite(tob_Champs.GetValue('ARS_UNITETEMPS'), gsUniteValo, 1);
      if (dCoeffConvert = 0) then
        dCoeffConvert := 1;
      dTauxUnit := tob_Champs.GetValue('ARS_TAUXUNIT') / dCoeffConvert;

      dCoeffConvert2 := AFConversionUnite(Article.GetValue('GA_QUALIFUNITEVTE'), gsUniteValo, 1);
      if (dCoeffConvert2 = 0) then
        dCoeffConvert2 := 1;
      dTauxUnitPres := Article.GetValue('GA_PMRP') / dCoeffConvert2;
    end;
  end;

  Result := 1;
  // ATTENTION, si modif de cette fct, faire idem dans TOM_RessourcePR (UtomRessource)
  AcsChampFormule := AnsiUppercase(AcsChampFormule);
  if (AcsChampFormule = 'TAUXUNITRES') then
    result := dTauxUnit
  else
    if (AcsChampFormule = 'FRAISGEN1') then
    result := tob_Champs.GetValue('ARS_TAUXFRAISGEN1')
  else
    if (AcsChampFormule = 'FRAISGEN2') then
    result := tob_Champs.GetValue('ARS_TAUXFRAISGEN2')
  else
    if (AcsChampFormule = 'COEFMETIER') then
    result := tob_Champs.GetValue('ARS_COEFMETIER')
  else
    if (AcsChampFormule = 'TAUXREVIENT') or (AcsChampFormule = 'TAUXVENTE') then
  begin
    if (ListePrix <> nil) then
    begin
      if (ListePrix.Count <> 0) then
        Result := Valeur(ListePrix[0]) / dCoeffConvert;
      ListePrix.Free;
    end;
  end
  else
    if (AcsChampFormule = 'TAUXPATRON') then
    result := tob_Champs.GetValue('ARS_TAUXCHARGEPAT')
  else
  begin
    // traitement des zones articles
    if ((AcsChampFormule = 'TAUXUNITPRES') or (AcsChampFormule = 'FRAISGENPRES')) then
    begin
      if (AcsChampFormule = 'TAUXUNITPRES') then
        result := dTauxUnitPres
      else
        if (Article <> nil) then
        if (AcsChampFormule = 'FRAISGENPRES') then
          result := Article.GetValue('GA_COEFFG');
    end;
  end;
end;

(******************************************************************************)
// PRRessource : renvoie le prix de revient de la ressource
//
//
(******************************************************************************)

function TAFO_Ressource.PRRessource(AcdDate: TDateTime; ActobArticle: TOB; ActobAffaires: TOB; AcsCodeArticle: string; AcsCodeAffaire: string = ''): double;
begin
  Result := CalculPrixRessource(AcdDate, 'R', ActobArticle, ActobAffaires, AcsCodeArticle, AcsCodeAffaire);
end;

(******************************************************************************)
// PVRessource : renvoie le prix de vente de la ressource
//
// Entrée : booleen indiquant si c'est le HT = True ou le TTC = False que l'on
// veut en sortie
//
(******************************************************************************)

function TAFO_Ressource.PVRessource(AcdDate: TDateTime; AcbHT: boolean; ActobArticle: TOB; ActobAffaires: TOB; AcsCodeArticle: string; AcsCodeAffaire: string = ''): double;
//var
//dCoeffConvert:double;
begin
  // on prend comme référence le PV HT ou TTC ?
  gbPVHT := AcbHT;

  Result := CalculPrixRessource(AcdDate, 'V', ActobArticle, ActobAffaires, AcsCodeArticle, AcsCodeAffaire);
  {Result := 0;
  if (gsUniteValo=tob_Champs.GetValue('ARS_UNITETEMPS')) then
     // Si l'unite de valorisation est la meme que l'unite de la ressource
     dCoeffConvert:=1
  else
     // Sinon on doit convertir en sortie
     begin
     dCoeffConvert := AFConversionUnite(tob_Champs.GetValue('ARS_UNITETEMPS'), gsUniteValo, 1);
     if dCoeffConvert=0 then dCoeffConvert :=  1;
     end;

  if (AcbHT = true) then
     // on veut le HT
     Result := tob_Champs.GetValue('ARS_PVHT') / dCoeffConvert
  else
     // on veut le TTC
     Result := tob_Champs.GetValue('ARS_PVTTC') / dCoeffConvert;
  }

end;

(******************************************************************************)
// RessourceDisponible : renvoie un booleen indiquant si la ressource est disponible
//                      à la date et heure entrée
//
// Entrée : Date et heure auxquelles on veut savoir si la ressource est disponible
//
// Il s'agit de la disponibilité par rapport aux dates et heures de présence
// (CDD, temps partiel, location...) normales de la ressource.
//
(******************************************************************************)

function TAFO_Ressource.RessourceDisponible(AcdDateHeureDemande: TDateTime): boolean;
var
  dDebutDispo: TDateTime;
  dFinDispo: TDateTime;
  sStandCalen: string;
  bCalenSpecif: boolean;
begin
  Result := False;

  dDebutDispo := tob_Champs.GetValue('ARS_DEBUTDISPO');
  dFinDispo := tob_Champs.GetValue('ARS_FINDISPO');
  sStandCalen := tob_Champs.GetValue('ARS_STANDCALEN');
  bCalenSpecif := tob_Champs.GetValue('ARS_CALENSPECIF');

  if (AcdDateHeureDemande < dDebutDispo)
    and (AcdDateHeureDemande >= dFinDispo) then
    // Si la date/heure fournie en entrée n'est pas dans l'intervalle de
    // disponibilité de la ressource, inutile d'aller voir plus en détail
    // la ressource n'est pas disponible
    exit
  else
    // en attendant les services de l'objet Gestion de calendrier, on renvoie vrai
    // si la date d'entrée et comprise dans les dates de la ressource
    Result := true;

  // On va regarder si la date/heure d'entrée est dans le calendrier lié à la ressource
  if (bCalenSpecif = false) then
    // Il ne s'agit pas d'un calendrier spécifique, on va lire le calendrier standard
    // dont le code est dans le champ STANDCALEN
  begin
  end
  else
    // On va lire le calendrier spécifique de la ressource
  begin
  end;
end;

(******************************************************************************)
//
// ListeCompetencesD1Ressource: renvoie une TStringList des compétences de la ressource
//
(******************************************************************************)

function TAFO_Ressource.ListeCompetencesD1Ressource: TStringList;
var
  sSelectR: string;
  QSelectR: TQuery;
  vslListCompet: TStringList;
  TobCompetRess: TOB;
  i: integer;
begin
  QSelectR := nil;
  Result := nil;

  sSelectR := 'SELECT ACR_COMPETENCE FROM COMPETRESSOURCE WHERE ACR_RESSOURCE="' + tob_Champs.GetValue('ARS_RESSOURCE') + '"';

  TobCompetRess := TOB.Create('', nil, -1);
  try
    QSelectR := OpenSQL(sSelectR, TRUE,-1,'',true);
    vslListCompet := TStringList.Create;

    TobCompetRess.LoadDetailDB('', '', '', QSelectR, true);
    for i := 0 to TobCompetRess.Detail.count - 1 do
    begin
      vslListCompet.Add(TobCompetRess.Detail[i].GetValue('ACR_RESSOURCE'));
    end;

  finally
    Ferme(QSelectR);
    TobCompetRess.Free;
  end;

  if (vslListCompet.Count <> 0) then
    Result := vslListCompet;

end;

end.
