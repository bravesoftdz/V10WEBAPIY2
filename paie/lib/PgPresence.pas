{***********UNITE*************************************************
Auteur  ...... : PAIE PGI :
Créé le ...... : 20/01/2007
Modifié le ... :   /  /
Description .. : Gestion des évènements de présence
Mots clefs ... : PAIE;ABSENCES
*****************************************************************}
{
PT1 30/10/2007 GGU V_80 Gestion des absences en horaire
PT2 26/11/2007 GGU V_80 Externalisation de la fonction EclateTob
PT3 07/02/2008 GGU V_81 Utilisation de 2 paramsoc pour gérer les dates de clôture
PT4 04/03/2008 NA  V_81 Correction Evenement plage 2
}
unit PgPresence;

interface

uses
  StdCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
  DB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HEnt1,UTOB,
  PGPOPULOUTILS, Classes;

const
  periodiciteAnnuelle     : String = '001';
  periodiciteMensuelle    : String = '002';
  periodiciteHebdomadaire : String = '003';
  periodiciteJournaliere  : String = '004';
  periodiciteFinDeCycle   : String = '005';

  // Types de périodes utilisées par les compteurs
  JOUR    = 1;
  SEMAINE = 2;
  MOIS    = 3;

  TYP_POPUL_PRES = 'PRE';

//procedure InitialiseTobPresenceSalarie(TCP: tob);
Function  RendLibPresence(TypeConge : string; DD,DF : TDateTime) : string;
Function  ControleGestionMaximumPresence(Salarie,TypePresence : String; T_MotifPersence : Tob; DateAbs : TDateTime; Nb_H : Double ) : Boolean;
Function  RecupMessageErrorPresence(Error : integer) : string;
Procedure GetDateHeureDebutFin(HORDEBPLAGE1, HORFINPLAGE1, HORDEBPLAGE2, HORFINPLAGE2 : TDateTime; JOURJ1PLAGE1, JOURJ1PLAGE2 : String; var Debut, Fin :TDateTime; DateCourante : TDateTime = 0);
Function GetParamSocDatesCloture(TypePeriode : ShortInt) : String;  //PT3
Procedure SetParamSocDatesCloture(TypePeriode : ShortInt; ListeParam : String); //PT3

{ Procedures communes de gestion des plannings de présence }
{ Crée un Item de planning représentant une journée type}
procedure CreateNewPresenceItem(JourneeType: String; DateCourante : TDateTime; RessourceID, RessourceLibelle: String; TobJourneeType, TobItems : Tob; IsException : Boolean = False; clef : String = '');

function GetAllValidesFromTobValidite(FromTob : Tob; NomChamps : Array of String;
           ValeurChamps : Array of Variant; NomChampDateValidite : String; NomChampDateDebut : String = 'DATEDEBUT'; DateDebut : TDateTime = 2;
           NomChampDateFin : String = 'DATEFIN'; LastDateFin : TDateTime = 73050): tob;
function GetAllValidesFromTobValiditeMultiPass(FromTob : Tob; NomChamps, NomChampsMultiPass : Array of String;
           ValeurChamps : Array of Variant; NomChampDateValidite : String; NomChampDateDebut : String = 'DATEDEBUT'; DateDebut : TDateTime = 2;
           NomChampDateFin : String = 'DATEFIN'; LastDateFin : TDateTime = 73050): tob;
Function GetLastValideFromTob(LaDate : TDateTime; ChampsDate : String; FromTob : Tob; NomChamps : Array of String; ValeurChamps : Array of Variant) : tob;

// Vérification de la validité d'un mouvement de présence
Function ControleMouvementsPresence (Action : String; T_MotifAbs : TOB; Salarie,TypeConges : String; DateDebAbs,DateFinAbs,HDeb,Hfin : TDateTime; NbHeures : Real; AbsKeyOrdre : Integer; AbsKeyRessource : String; DateEntree : TDateTime = 2; DateSortie : TDateTime = 2) : ShortInt;

  // Positionne un compteur à recalculer pour un ou tous les salariés à la date donnée
  Procedure CompteursARecalculer(DateRecalcule : TDateTime; Salarie : String = 'TOUS');

  // Récupération du profil de présence d'un salarié en fonction de sa population de rattachement
  //function GetProfilPresenceFromPop (Salarie : String; LaDate : TDateTime=0; GP : TGestionPresence=Nil; Keep : Boolean=False) : String;

  // Fonction permettant de savoir si les éléments sont utilisés (pour vérification d'impacts)
  Function PresenceCycleIsUsed (TypeAffect, CycleAffect : String) : Boolean;  // Cycles
  Function PresenceMeterIsUsed  (Compteur : String) : Boolean;                // Compteurs
  Function PresenceProfileIsUsed (Profile : String) : Boolean;                // Profil
  Function PresenceJourneyIsUsed (Journey : String) : Boolean;                // Journée

  // Fonctions de récupération et de sauvegarde des dates traitées par le moteur de calcul
  Function  PresenceRenvoieParamSoc  (TypePeriode : ShortInt) : String;
  Procedure PresenceSauveJourCalcule (DateJour : TDateTime);
//  Procedure PresenceSauveMoisCalcule (Fin : TDateTime; Recalcul : Boolean = False);
//  Procedure PresenceSauvePeriodeCalculee (TypePeriode : ShortInt; Fin : TDateTime; Reinit : Boolean = False; Intel : Boolean = True);
  Procedure PresenceDonneMoisCalculActuel(var Debut, Fin : TDateTime);
  Procedure PresenceDonneSemaineCalculActuelle(var Debut, Fin : TDateTime);
  Procedure PresenceDonnePeriodeCalcul (TypePeriode,NumPeriode : ShortInt; var Debut, Fin : TDateTime);

type
  TModeleAndJournee = class
  public
    Modele, Journee : String;
    DateLancementCycle, DateDebutModele, DateFinModele : TDateTime;
    resteJours : Integer;
  end;

{ Objet de gestion de la présence }
type
  TGestionPresence = class
  private
    InMemoryBeginDate, InMemoryEndDate : TDateTime;
    FFiltreSalarie, FPSAFiltreSalarie : String;
    SavTobModeleAndJournee : Tob;
    Function    GetException(LaDate : TDateTime; Cycle : String; TypeDeCycle : String = 'CYC') : String;
    Function    GetExceptionSalarie(LaDate : TDateTime; Salarie : String; var CycleOuModele, TypeDeCycle : String) : Boolean;
//    procedure   GetSavedModeleAndJourneeInCycle(var Modele, Journee : String;
//      var DateLancementCycle, DateDebutModele, DateFinModele : TDateTime;
//      var resteJours : Integer; LaDate : TDateTime; Cycle : String;
//      TypeDeCycle : String = 'CYC'; CheckException : Boolean = True) ;
  protected
    function    GetFiltreSalarie(const Prefixe : String): String;
    procedure   SetFiltreSalarie(const Prefixe, Value: String);
  public
    TobCycle, TobModeleCycle, TobAssocCycle, TobAssocModele, TobDureeCycle, TobDureeCycleModele, TobExceptions : Tob; // , TobJourTypes
    TobProfilPresenceSalarie, TobProfilPresence, TobExceptionSalaries, TobSalariePopul : Tob;
    TobParametresPopul, TobParametresAssocPopul, TobAbsencePresenceSalarie : Tob;
    constructor Create( NeedSalarie : Boolean = True; NeedPopulation : Boolean = True; NeedException : Boolean = True; NeedCycles : Boolean = True; NeedAbsencePresence : Boolean = True; Du : TDateTime = 0; Au : TDateTime = 0; OnlyPop : Boolean=False; Salarie : String='');
    destructor  Destroy; override;
//PT2    procedure EclateTob(var FromTob, ToTob: Tob; ChampCritere: String; Duplicate : Boolean = True); // ; NoDoublons : Boolean = False
    Procedure AddFiltreSalarie(const Prefixe : String; var StQuery : String);
    procedure GetModeleAndJourneeInCycle(var Modele, Journee : String;
      var DateLancementCycle, DateDebutModele, DateFinModele : TDateTime;
      var resteJours : Integer; LaDate : TDateTime; Cycle : String;
      TypeDeCycle : String = 'CYC'; CheckException : Boolean = True) ;
    procedure GetDatesModeleCycle(var DateDebut, DateFin: TDateTime;
      LaDate: TDateTime; Cycle : String; TypeDeCycle : String = 'CYC');
    function  GetTobAbsencePresenceFiltreSalarie(Salarie: string): Tob;
    Function  GetJourneeType(LaDate : TDateTime; Cycle : String; TypeDeCycle : String = 'CYC'; CheckException : Boolean = True) : String;
    Function  GetJourneeTypeSalarie(LaDate : TDateTime; Salarie : String; var CycleOuModele, TypeDeCycle : String; var isExceptionSalarie, isExceptionCycle : Boolean; CheckException : Boolean = True) : String;
    Procedure GetInfoPresenceSalarie(LaDate : TDateTime; Salarie : String; var Profil, Cycle, Modele, JourneeType : String; var isExceptionSalarie, isExceptionCycle : Boolean; CheckException : Boolean = True);
    Function  GetTobExceptionSalarie(LaDate : TDateTime; Salarie : String) : Tob;
    Function  GetTobAbsencePresenceSalarie(LaDate : TDateTime; Salarie : String; FiltreTypeMvt : String = '') : Tob;
    Function  GetAbsenceSalarie (LaDate: TDateTime; Salarie: String; var Ressource : String; var Ordre : Integer): String;
    Function  GetPresenceSalarie(LaDate: TDateTime; Salarie: String; var Ressource : String; var Ordre : Integer): String;
    Function  GetAbsencePresenceSalarie(LaDate: TDateTime; Salarie: String; var TypeMvt, Ressource : String; var Ordre : Integer): String;
    function  GetProfilsPresenceSalarie(Salarie: String): tob;
    Function  GetProfil(Salarie : String; LaDate : TDateTime; TypeProfilPres : String='' ) : String;
    Function  UpdateProfil(Salarie : String; LaDate : TDateTime) : Boolean;
    function  GetTypesEtProfilsPresenceSalarie(Salarie: String; TobProfilsPresenceSalarie: tob): tob;
    function  GetRythmesSalarie(TobTypesEtProfilsPresenceSalarie: tob): tob;  //Salarie: String;
    function  GetCyclesSalarie(TobRythmesSalarie: tob): tob; //Salarie: String;
    function  GetDecompositionModelesCyclesSalarie(TobCyclesSalarie: tob; DateDebutMin, DateFinMax : TDateTime): tob; //Salarie: String;
    Procedure ReloadAssocModele;
    Procedure ReloadAssocCycle;
    Procedure ReloadDureeCycleModele;
    Procedure ReloadProfilPresenceSalarie;
    Procedure ReloadProfilPresence;
    Procedure ReloadDureeCycle;
    Procedure ReloadSalariePopul (Salarie : String='');
    Procedure ReloadParametresAssocPopul;
    Procedure ReloadParametresPopul;
    Procedure ReloadException;
    Procedure ReloadExceptionSal;
    Procedure ReloadAbsencePresenceSalarie(Du, Au : TDateTime; Append : Boolean = True; FiltreSalarie : String = '');
    Function  DateOfException(LaDate : TDateTime; Cycle : String; TypeDeCycle : String = 'CYC') : Boolean;
  end;

  TCalcuVarPre = class(TGestionPresence)
  private
    { Liste des compteurs utilisés par les variables de présence en fonction de leur périodicité }
    ListeCompteurs412, ListeCompteurs813, ListeCompteurs9, ListeCompteurs1014, ListeCompteurs11 : String;
    TobDroitsJourneesTypes, TobMotifsEvenements : Tob;
    idDATEDEBUT, idDATEFIN, idPYR_PERIODICITEPRE : Integer;  //idPYR_PERIODERAZ  { Numéro de champs pour la fonction de traitement de la périodicité }
    IsNumChampSaved : Boolean;
  public
    Tob_Ferie : Tob;
    TobValeursCompteurs, TobJourneesTypes, TobCompteurs : Tob;
    TobSaveCalcVar : Tob;  { Tob permettant de sauvegarder les variables calculées au fur et a mesure du traitement, ça évite les recalculs... }
    ListeFinDeMois : Array of TDateTime; //ListeFinDeMois[0] = Fin du mois en cours, ListeFinDeMois[1] = Fin du mois -1 ...
    CalculJour, CalculHebdo, CalculFinCycle, CalculMens, CalculAnn : Boolean;
    constructor Create(FiltreSalarie : String = ''; DateDebut : TDateTime = 0; DateFin : TDateTime = 0; DateFinEstFinDeMois : Boolean = False); overload;
    destructor Destroy; override;
    function BuildLoadingQuery(DateDebut, DateFin : TDateTime) : String;
    function GetTobValeursCompteursSalarie(Salarie : string) : Tob;
    function EvalVarPreEvt(Salarie,MotifEvenement : string; const ResultColonneName : String; const DateDeb, DateFin: TDateTime; Diag: TObject = nil) : double;
    function EvalVarPreCompPres(Salarie,CompteurPresence: string; Perio : String;
      const DateDeb, DateFin: TDateTime; TOB_Rub : Tob; var NeedSave, HadRegul : Boolean; var Indicateur : String; Diag: TObject = nil; ForceRecalcul : Boolean = False) : double;
    function EvalVarPreJType(Salarie,Droit: string; const DateDeb,
      DateFin: TDateTime; TypeBase : Integer; Diag: TObject = nil; JourneeType : String = 'A Rechercher'): double;
    procedure ConcatRegul(Salarie, CompteurPresence, Perio: String;
      const DateDeb, DateFin: TDateTime);
    procedure Traiteperiodicite(var TobCompteursACalculer: tob; Salarie : String; DateEntreeSalarie, DateSortieSalarie : TDateTime;
//      CalculJour, CalculHebdo, CalculMens, CalculAnn, CalculFinCycle: Boolean;
      DateDebutJour, DateDebutHebdo, DateDebutMens, DateDebutAnn, DateDebutFinCycle,
      DateFinJour, DateFinHebdo, DateFinMens, DateFinAnn, DateFinFinCycle : TDateTime);
    Function Str5cDateToStr10cDate(Const RefDate : TDateTime; strDate5c: String) : String;
    procedure GetCurrentMonth(var DateDebut, DateFin : TDateTime);
    procedure GetMonth(const Value : TDateTime; var DateDebut, DateFin : TDateTime); overload;
    procedure GetMonth(const decalage : Integer; const reference : TDateTime; var DateDebut, DateFin : TDateTime); overload;
    procedure GetMonthAbs(const decalage : Integer; var DateDebut, DateFin : TDateTime); overload;
    procedure GetCurrentWeek(var DateDebut, DateFin : TDateTime);
    procedure GetWeek(const Value : TDateTime; var DateDebut, DateFin : TDateTime); overload;
    procedure GetWeek(const decalage : Integer; var DateDebut, DateFin : TDateTime; const reference : TDateTime = 0); overload;
    Procedure FreeMemSalarie(Salarie : String);
  end;

  // Récupération du profil de présence d'un salarié en fonction de sa population de rattachement
  function GetProfilPresenceFromPop (Salarie : String; LaDate : TDateTime=0) : String;
  // Mise à jour du profil de présence d'un salarié en fonction de sa population de rattachement
  procedure UpdateProfilPresenceFromPop (Salarie : String; LaDate : TDateTime=0);

implementation

uses

  hctrls, SysUtils, DateUtils, HeureUtil, Math, PGCalendrier, Paramsoc, P5Util, StrUtils
  ,PGTobOutils //PT2
  ;

Function GetParamSocDatesCloture(TypePeriode : ShortInt) : String; //PT3
var
  NomParamSoc : String;
begin
  NomParamSoc := PresenceRenvoieParamSoc(TypePeriode);
  result := GetParamSocSecur(NomParamSoc, '');
  if NomParamSoc = 'SO_PGPRESMOISTRAITES' then
  begin
   result := result + GetParamSocSecur(PresenceRenvoieParamSoc(TypePeriode)+'1', '');
   result := result + GetParamSocSecur(PresenceRenvoieParamSoc(TypePeriode)+'2', '');
  end;
end;

Procedure SetParamSocDatesCloture(TypePeriode : ShortInt; ListeParam : String);  //PT3
var
  NomParamSoc : String;
begin
  NomParamSoc := PresenceRenvoieParamSoc(TypePeriode);
  SetParamSoc(NomParamSoc, LeftStr(ListeParam, 70));
  if NomParamSoc = 'SO_PGPRESMOISTRAITES' then
  begin
    SetParamSoc(NomParamSoc+'1', Copy(ListeParam, 71, 70) );
    SetParamSoc(NomParamSoc+'2', Copy(ListeParam, 141, 70) );
  end;
end;

Function PremierJourSemaineEstLundi : Boolean;
var
  Param : String;
begin
  { Rechercher le parametre société précisant si le jour de la semaine est un lundi ou un dimache }
  Param := GetParamSocSecur('SO_PGPRESJOURDEBSEMAINE', '1LU');
  result := (Param = '1LU');
end;

Function JourDeLaSemaine(const AValue : TDateTime) : Word;
begin
  if PremierJourSemaineEstLundi then
    result := DayOfTheWeek(AValue)
  else
    result := DayOfWeek(AValue);
end;

Function SemaineDeLAnnee(const AValue : TDateTime) : Word;
begin
  if PremierJourSemaineEstLundi or (DayOfTheWeek(AValue) <> DaySunday) then
    result := WeekOfTheYear(AValue)
  else begin
    result := WeekOfTheYear(AValue)+1;
    if result > WeeksInYear(AValue) then result := 1;
  end;
end;

Function DebutDeLaSemaine(const AValue: TDateTime): TDateTime;
begin
  result := AValue - JourDeLaSemaine(AValue) + 1;
end;

Function FinDeLaSemaine(const AValue: TDateTime): TDateTime;
begin
  result := DebutDeLaSemaine(AValue) + 6;
end;

Procedure GetDateHeureDebutFin(HORDEBPLAGE1, HORFINPLAGE1, HORDEBPLAGE2, HORFINPLAGE2 : TDateTime; JOURJ1PLAGE1, JOURJ1PLAGE2 : String; var Debut, Fin :TDateTime; DateCourante : TDateTime = 0);
var
  WYear1, WMonth1, WDay1, WHour1, WMinute1, WSeconde1, WMilli1 : Word;
  WYear2, WMonth2, WDay2, WHour2, WMinute2, WSeconde2, WMilli2 : Word;
begin
  if DateCourante < 10 then DateCourante := Date;
  DecodeDateTime(HORDEBPLAGE1, WYear1, WMonth1, WDay1, WHour1, WMinute1, WSeconde1, WMilli1);
  DecodeDateTime(DateCourante, WYear2, WMonth2, WDay2, WHour2, WMinute2, WSeconde2, WMilli2);
  Debut := EncodeDateTime(WYear2, WMonth2, WDay2, WHour1, WMinute1, WSeconde1, WMilli1);

  DecodeDateTime(HORFINPLAGE2, WYear1, WMonth1, WDay1, WHour1, WMinute1, WSeconde1, WMilli1);
  DecodeDateTime(HORDEBPLAGE2, WYear2, WMonth2, WDay2, WHour2, WMinute2, WSeconde2, WMilli2);
  if (WHour1 = WHour2) and (WMinute1 = WMinute2) then
  begin  //Il n'y a pas de deuxième plage horaire, on prends la fin de la plage 1
    DecodeDateTime(HORFINPLAGE1, WYear1, WMonth1, WDay1, WHour1, WMinute1, WSeconde1, WMilli1);
    if (JOURJ1PLAGE1 <> 'X') then
      DecodeDateTime(DateCourante,WYear2, WMonth2, WDay2, WHour2, WMinute2, WSeconde2, WMilli2)
    else
      DecodeDateTime((DateCourante+1),WYear2, WMonth2, WDay2, WHour2, WMinute2, WSeconde2, WMilli2);
  end else begin
    if (JOURJ1PLAGE2 <> 'X') then
      DecodeDateTime(DateCourante,WYear2, WMonth2, WDay2, WHour2, WMinute2, WSeconde2, WMilli2)
    else
      DecodeDateTime((DateCourante+1),WYear2, WMonth2, WDay2, WHour2, WMinute2, WSeconde2, WMilli2);
  end;
  Fin := EncodeDateTime(WYear2, WMonth2, WDay2, WHour1, WMinute1, WSeconde1, WMilli1);
end;


Function GetLastValideFromTob(LaDate : TDateTime; ChampsDate : String; FromTob : Tob; NomChamps : Array of String; ValeurChamps : Array of Variant) : tob;
var
  LastTestDate, TestDate : TDateTime;
  TempTob : Tob;
begin
  result := nil;
  LastTestDate := iDate1900;
  TempTob := FromTob.FindFirst(NomChamps, ValeurChamps, False);
  While Assigned(TempTob) do
  begin
    TestDate := TempTob.GetDateTime(ChampsDate);
    { On recherche le profil le plus rescent mais inférieur à la date de recherche }
    if  (LaDate >= TestDate) and (LastTestDate < TestDate) then
    begin
      result := TempTob;
      LastTestDate := TestDate;
    end;
    TempTob := FromTob.FindNext(NomChamps, ValeurChamps, False);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GG / FLO
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : Vérifie la validité d'un mouvement de présence
Retours ...... : 15 : Une absence existe déjà sur la période
               : 16 : Salarié sorti à la date du mouvement
               : 17 : Quantité maximum dépassée
               : 18 : Nb heures maximum dépassé
               : 23 : Salarié non encore entré à la date du mouvement
               : 32 : Une présence existe déjà sur la période
               Mots clefs ... :
*****************************************************************}
function ControleMouvementsPresence (Action : String; T_MotifAbs : TOB; Salarie,TypeConges : String;
    DateDebAbs,DateFinAbs,HDeb,Hfin : TDateTime; NbHeures : Real; AbsKeyOrdre : Integer; AbsKeyRessource : String; DateEntree : TDateTime = 2; DateSortie : TDateTime = 2) : ShortInt;
var
  Q: TQuery;
  st: string;
  DateD, DateF : tDateTime;
  NbPresenceSurPeriode : Integer;
  HDebutMvt, MDebutMvt, HFinMvt, MFinMvt, HDebutMvtTrouve, MDebutMvtTrouve, HFinMvtTrouve, MFinMvtTrouve, Sec, MSec : Word;
  HeurD, HeurF  : tDateTime;
  HeureDebutMvt, HeureFinMvt, HeureDebutMvtTrouve, HeureFinMvtTrouve: TDateTime;
begin
     Result := 0;

     if Assigned(T_MotifAbs) then
     begin
          // Teste le respect de la gestion des maximums
          if (T_MotifAbs.GetValue('PMA_GESTIONMAXI') = 'X') then
          begin
               If NbHeures > 0 Then
               begin
                    If ControleGestionMaximumPresence(Salarie,TypeConges,T_MotifAbs,DateDebAbs,NbHeures) then
                    Begin
                         if T_MotifAbs.GetValue('PMA_JOURHEURE') = 'QTE' then
                         begin
                              Result := 17;
                              exit;
                         end
                         else if (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU') or (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HOR') then //PT1
                         begin
                              Result := 18;
                              exit;
                         End;
                    End;
               End;
          End;

          if (T_MotifAbs.GetValue('PMA_CONTROLMOTIF') = 'X') then
          begin
               DateD := DateDebAbs;
               DateF := DateFinAbs;

               // Si la date d'entrée n'est pas renseignée, c'est que les paramètres ne sont pas passés
               // On va donc les rechercher en table
               If (DateEntree = iDate1900) Then
               Begin
                    // Vérification de la date de sortie du salarié
                    St := 'SELECT PSA_SALARIE FROM SALARIES ' +
                          'WHERE PSA_SALARIE  ="' + Salarie               + '" ' +
                          'AND ( PSA_DATESORTIE >="' + UsDateTime(DateF)     + '" ' +
                          '   OR PSA_DATESORTIE ="' + UsdateTime(Idate1900) + '" ' +
                          '   OR PSA_DATESORTIE IS NULL )';
                    Q := opensql(st, True);
                    if Q.eof then
                    begin
                         Result := 16;
                         Ferme(Q);
                         exit;
                    end;
                    ferme(Q);

                    // Contrôle de la date d'entrée du salarié
                    St := 'SELECT PSA_SALARIE FROM SALARIES ' +
                          'WHERE PSA_SALARIE  = "' + Salarie          + '" ' +
                          ' AND PSA_DATEENTREE <="' + usdatetime(DateD)+ '" ';
                    Q := opensql(st, True);
                    if Q.eof then
                    begin
                         Result := 23;
                         Ferme(Q);
                         exit;
                    end;
                    ferme(Q);
               End
               Else
               Begin
                    // Vérification de la date de sortie du salarié
                    If (DateSortie>iDate1900) And (DateSortie < DateF) Then
                    Begin
                         Result := 16;
                         Exit;
                    End;

                    // Contrôle de la date d'entrée du salarié
                    If DateEntree > DateD Then
                    Begin
                         Result := 23;
                         Exit;
                    End;
               End;
          end;

          // On teste si une absence existe sur la meme période
          if (T_MotifAbs.GetValue('PMA_CTRLABSEXISTE') = 'X') then
          begin
               DateD := DateDebAbs;
               DateF := DateFinAbs;

               St := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_TYPECONGE,PCN_TYPEMVT,PCN_ORDRE FROM ABSENCESALARIE ' +
                     'WHERE PCN_SALARIE = "' + Salarie + '" ' +
                     'AND PCN_TYPEMVT <> "PRE" ' +
                     'AND PCN_ETATPOSTPAIE <> "NAN" ' + //Absences non annulées
                     'AND (   ( PCN_DATEDEBUTABS >"' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS < "' + usdatetime(DateF) + '") '+
                     '     OR ( PCN_DATEFINABS   >"' + usdatetime(DateD) + '" AND PCN_DATEFINABS   < "' + usdatetime(DateF) + '") '+
                     '     OR ( PCN_DATEDEBUTABS <"' + usdatetime(DateD) + '" AND PCN_DATEFINABS   > "' + usdatetime(DateF) + '"))';
               Q := opensql(st, True);
               if not Q.eof then
               begin
                    if ((Action = 'CREATION') and (Q.RecordCount = 1))
                    or ((Action = 'MODIFICATION') and (Q.RecordCount > 1)) then
                    begin
                         Result := 15;
                         Ferme(Q);
                         exit;
                    end;
               end;
               ferme(Q);
          end;

          // Contrôle de chevauchement d'évènement de présence
          if (T_MotifAbs.GetValue('PMA_CTRLPREEXISTE') = 'X') then
          begin
               DateD := DateDebAbs;
               DateF := DateFinAbs;
               HeurD := HDeb;
               HeurF := HFin;
               DecodeTime(HeurD, HDebutMvt, MDebutMvt, Sec, MSec);
               HeureDebutMvt := EncodeTime(HDebutMvt, MDebutMvt, Sec, MSec);
               DecodeTime(HeurF, HFinMvt, MFinMvt, Sec, MSec);
               HeureFinMvt := EncodeTime(HFinMvt, MFinMvt, Sec, MSec);
               NbPresenceSurPeriode := 0;

               // Chevauchements
               St := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS ' +
                     'FROM ABSENCESALARIE ' +
                     'WHERE PCN_SALARIE = "' + Salarie + '" ' +
                     'AND PCN_TYPEMVT = "PRE" ' +
                     'AND (   ( PCN_DATEDEBUTABS > "' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS < "' + usdatetime(DateF) + '") '+
                     '     OR ( PCN_DATEFINABS   > "' + usdatetime(DateD) + '" AND PCN_DATEFINABS   < "' + usdatetime(DateF) + '") '+
                     '     OR ( PCN_DATEDEBUTABS < "' + usdatetime(DateD) + '" AND PCN_DATEFINABS   > "' + usdatetime(DateF) + '") '+
                     '    )'+
                     'AND NOT (PCN_TYPEMVT = "PRE" and PCN_SALARIE = "' + Salarie + '" and PCN_ORDRE = ' + IntToStr(AbsKeyOrdre) + ' and PCN_RESSOURCE = "' + AbsKeyRessource + '")';

               Q := opensql(st, True);
               if not Q.eof then
               begin
                    NbPresenceSurPeriode := Q.RecordCount;
               end;
               Ferme(Q);

               // Résolution des cas limites (Periode commencant ou finissant le meme jour
               // que la periode a créer) par le test des heures
               St := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_HDEB,PCN_HFIN ' +
                     'FROM ABSENCESALARIE ' +
                     'WHERE PCN_SALARIE = "' + Salarie + '" ' +
                     'AND PCN_TYPEMVT = "PRE" ' +
                     'AND (   ( PCN_DATEFINABS   ="' + usdatetime(DateD) + '") '+
                     '     OR ( PCN_DATEDEBUTABS ="' + usdatetime(DateF) + '") '+
                     '    )'+
                     'AND NOT (PCN_TYPEMVT = "PRE" and PCN_SALARIE = "' + Salarie + '" and PCN_ORDRE = ' + IntToStr(AbsKeyOrdre) + ' and PCN_RESSOURCE = "' + AbsKeyRessource + '")';
               Q := opensql(st, True);
               while not Q.eof do
               begin
                    DecodeTime((Q.FindField('PCN_HDEB').AsDateTime), HDebutMvtTrouve, MDebutMvtTrouve, Sec, MSec);
                    HeureDebutMvtTrouve := EncodeTime(HDebutMvtTrouve, MDebutMvtTrouve, Sec, MSec);
                    DecodeTime((Q.FindField('PCN_HFIN').AsDateTime), HFinMvtTrouve, MFinMvtTrouve, Sec, MSec);
                    HeureFinMvtTrouve := EncodeTime(HFinMvtTrouve, MFinMvtTrouve, Sec, MSec);

                    If (Q.FindField('PCN_DATEFINABS').AsDateTime = DateF) And (Q.FindField('PCN_DATEDEBUTABS').AsDateTime = DateD) Then
                    Begin
                      If( ((HeureDebutMvt >= HeureDebutMvtTrouve) And (HeureDebutMvt <= HeureFinMvtTrouve)) Or
                         ((HeureFinMvt   >= HeureDebutMvtTrouve) And (HeureFinMvt <= HeureFinMvtTrouve)) Or
                         ((HeureDebutMvt <= HeureDebutMvtTrouve) And (HeureFinMvt >= HeureFinMvtTrouve)) Or
                         ((HeureDebutMvt >= HeureDebutMvtTrouve) And (HeureFinMvt <= HeureFinMvtTrouve)) ) Then
                         Inc(NbPresenceSurPeriode);
                    End
                    else if ((HeureFinMvtTrouve >= HeureDebutMvt) and (Q.FindField('PCN_DATEFINABS').AsDateTime = DateD))
                         or ((HeureDebutMvtTrouve <= HeureFinMvt) and (Q.FindField('PCN_DATEDEBUTABS').AsDateTime = DateF)) then
                    begin
                         Inc(NbPresenceSurPeriode);
                    end;
                    Q.Next;
               end;
               if NbPresenceSurPeriode >= 1 then
               begin
                    Result := 32;
                    Ferme(Q);
                    exit;
               end;
               ferme(Q);
          end;
     end;
end;


function GetAllValidesFromTobValidite(FromTob : Tob; NomChamps : Array of String;
           ValeurChamps : Array of Variant; NomChampDateValidite : String; NomChampDateDebut : String = 'DATEDEBUT'; DateDebut : TDateTime = 2;
           NomChampDateFin : String = 'DATEFIN'; LastDateFin : TDateTime = 73050): tob;
var
  NewtobLigne, FindtobLigne, toblignes : Tob;
  indexTob : Integer;
  CalcDateFin : TDateTime;
begin
  toblignes := TOB.Create(FromTob.NomTable, nil, -1);
  result := toblignes;
  { Recherche des lignes vérifiant les conditions }
  FindtobLigne := FromTob.FindFirst(NomChamps, ValeurChamps, False);
  while Assigned(FindtobLigne)
    and (FindtobLigne.GetDateTime(NomChampDateValidite) <= LastDateFin) do
  begin
    { Ajout du champs dateFin pour la ligne précédente }
    if toblignes.FillesCount(0)>=1 then
    begin
      CalcDateFin :=  FindtobLigne.GetDateTime(NomChampDateValidite)-1 ;
      toblignes.Detail[toblignes.FillesCount(0)-1].AddChampSupValeur(NomChampDateFin, CalcDateFin);
    end;
    NewtobLigne := TOB.Create(toblignes.NomTable, toblignes, -1);
    NewtobLigne.Dupliquer(FindtobLigne, True, True);
    NewtobLigne.AddChampSupValeur(NomChampDateDebut,FindtobLigne.GetDateTime(NomChampDateValidite));
    if NewtobLigne.GetDateTime(NomChampDateDebut) < DateDebut then
      NewtobLigne.PutValue(NomChampDateDebut, DateDebut);
    FindtobLigne := FromTob.FindNext(NomChamps, ValeurChamps, False);
  end;
  { Ajout du champs dateFin pour la dernière ligne }
  if toblignes.FillesCount(0)>=1 then
    toblignes.Detail[toblignes.FillesCount(0)-1].AddChampSupValeur(NomChampDateFin, LastDateFin);
  { Suppression des details dont la date de début est supperieur à la date de fin }
  for indexTob := toblignes.FillesCount(0)-1 downto 0 do
  begin
    FindtobLigne := toblignes.Detail[indexTob];
    if FindtobLigne.GetDateTime(NomChampDateDebut) > FindtobLigne.GetDateTime(NomChampDateFin) then
    begin
      FindtobLigne.ChangeParent(nil,-1);
      FreeAndNil(FindtobLigne);
    end;
  end;
end;


function GetAllValidesFromTobValiditeMultiPass(FromTob : Tob; NomChamps, NomChampsMultiPass : Array of String;
           ValeurChamps : Array of Variant; NomChampDateValidite : String; NomChampDateDebut : String = 'DATEDEBUT'; DateDebut : TDateTime = 2;
           NomChampDateFin : String = 'DATEFIN'; LastDateFin : TDateTime = 73050): tob;
var
  toblignes, FindtobLigne, TempTob, TempTobMere : Tob;
  indexKey, indexPartialResult : Integer;
  StTempKey : String;
  ListKeyTob : TStringList;
begin
  toblignes := TOB.Create(FromTob.NomTable, nil, -1);
  result := toblignes;
  ListKeyTob := TStringList.Create;
  { Recherche des lignes vérifiant les conditions }
  FindtobLigne := FromTob.FindFirst(NomChamps, ValeurChamps, False);
  while Assigned(FindtobLigne) do
  begin { Tri de ces lignes en fonction de leur clef }
    { Récupération de la valeur de la clef }
    StTempKey := '';
    for indexKey := 0 to Length(NomChampsMultiPass) -1 do
    begin
      StTempKey := StTempKey + '¤' + FindtobLigne.GetString(NomChampsMultiPass[indexKey]);
    end;
    { Si la même clef a déjà été trouvé, on récupère le groupe, sinon on en cré un nouveau }
    if ListKeyTob.Find(StTempKey, indexKey) then
    begin
      TempTobMere := ListKeyTob.Objects[indexKey] as TOB;
    end else begin
      TempTobMere := tob.Create('Groupe de :'+StTempKey, nil, -1);
      ListKeyTob.AddObject(StTempKey, TempTobMere);
    end;
    { Affectation de la ligne à son groupe }
    TempTob := tob.Create(toblignes.NomTable, TempTobMere, -1);
    TempTob.Dupliquer(FindtobLigne, True, True);
    FindtobLigne := FromTob.FindNext(NomChamps, ValeurChamps, False);
  end;
  { Appel de la fonction de validation pour chaque groupe }
  for indexKey := 0 to ListKeyTob.Count -1 do
  begin
    TempTobMere := ListKeyTob.Objects[indexKey] as TOB;
    FindtobLigne := GetAllValidesFromTobValidite(TempTobMere, NomChamps, ValeurChamps, NomChampDateValidite,
                                                 NomChampDateDebut, DateDebut, NomChampDateFin, LastDateFin);
    { Affectation du resultat partiel au résultat global }
    for indexPartialResult := FindtobLigne.FillesCount(0) - 1 downto 0 do
    begin
      FindtobLigne.Detail[indexPartialResult].ChangeParent(toblignes, -1);
    end;
    { Destruction de la mère partielle }
    FreeAndNil(FindtobLigne);
  end;
  { Destruction des groupes }
  for indexKey := 0 to ListKeyTob.Count -1 do
  begin
    TempTobMere := ListKeyTob.Objects[indexKey] as TOB;
    FreeAndNil(TempTobMere);
  end;
  FreeAndNil(ListKeyTob);
end;

Function RendLibPresence(TypeConge : string; DD,DF : TDateTime) : string;
var MonFormat : String;
Begin
  if TypeConge <> '' then Result := RechDom('PGMOTIFPRESENCE', Typeconge, False)
  else Result := '';
 if (DD <> idate1900) and (DF <> idate1900) and (Result <> '') then
 begin
    if Length(Result) > 14 then Result := Copy(Result, 1, 14); //Tronquage du libellé : recup des 14 premiers caractères
    If V_PGI.LanguePrinc = 'UK' then MonFormat := 'mmm dd yyyy'
      else MonFormat := 'dd/mm/yy';
    Result := TraduireMemoire(Result) + ' ' + FormatDateTime(MonFormat, DD) + ' '+TraduireMemoire('au')+' '+ FormatDateTime(MonFormat, DF);
    Result := Trim(Result);
 end;
End;

Function ControleGestionMaximumPresence(Salarie,TypePresence : String; T_MotifPersence : Tob; DateAbs : TDateTime; Nb_H : Double ) : Boolean;
var
  Per : Integer;
  DateDebPer : TDateTime;
  StSql, StType : String;
  Q : TQuery;
  YY,MM,JJ : WORD;
  Duree : Double ;
Begin
  Result := False;
  if T_MotifPersence.GetValue('PMA_GESTIONMAXI') = 'X' then
    begin
      StType := 'PCN_HEURES';
      Duree := Nb_H;

      if T_MotifPersence.GetValue('PMA_JRSMAXI') < Duree then
        Begin
        Result := True;
        End
      else
      if (T_MotifPersence.GetValue('PMA_TYPEPERMAXI')<>'') AND (T_MotifPersence.GetValue('PMA_PERMAXI') > 0) then
        Begin
        Per := T_MotifPersence.GetValue('PMA_PERMAXI');
        DateDebPer := IDate1900;
        DecodeDate(DateAbs,YY,MM,JJ);
        if T_MotifPersence.GetValue('PMA_TYPEPERMAXI')='PER' then
          Begin
          if MM < Per then YY := YY - 1;
          DateDebPer := EncodeDate(YY,Per,1);
          End
        else
          if T_MotifPersence.GetValue('PMA_TYPEPERMAXI')='GLI' then
             Begin
             DateDebPer := PlusDate(DateAbs,-Per,'M');
             End;
        StSql := 'SELECT SUM('+StType+') DUREE FROM ABSENCESALARIE '+
                 'WHERE PCN_TYPEMVT="PRE" AND PCN_SALARIE="'+Salarie+'" AND PCN_TYPECONGE="'+TypePresence+'" '+
                 'AND PCN_DATEDEBUTABS>="'+USDateTime(DateDebPer)+'" '+
                 'AND PCN_DATEFINABS<"'+USDateTime(DateAbs)+'" AND PCN_ETATPOSTPAIE <> "NAN" '; { PT12 }
        Q := OpenSql(StSql,True);
        If Not Q.Eof then
          Begin
          if T_MotifPersence.GetValue('PMA_JRSMAXI') < (Q.FindField('DUREE').AsFloat + Duree) then
            result := True;
          End;
        Ferme(Q);
        End;
    End; 
End;

Function  RecupMessageErrorPresence(Error : integer) : string;
Begin
  Case Error of
    1  : Result := 'La date de fin ne peut être inférieure à la date de début du mouvement.';
    4  : Result := 'Vous devez renseigner le type du mouvement.';
    9  : Result := 'Vous devez renseigner le type de mouvement.';
    10 : Result := 'Un mouvement ne peut être à cheval sur la date de clôture.';
    11 : Result := 'Vous devez renseigner la date de début.';
    12 : Result := 'Vous devez renseigner la date de fin.';
    14 : Result := ''; {Juste pour bloquer la validation}
    15 : Result := 'Un mouvement d''absence existe déjà sur cette période. Vous devez modifier vos dates de mouvement.';
    16 : Result := 'La date de fin du mouvement doit être antérieure à la date de sortie du salarié.';
    17 : Result := 'La quantité maximum octroyée pour ce motif est dépassée : ' {ComplMessage};
    18 : Result := 'Le nombre d''heures maximum octroyé pour ce motif est dépassé : ' {ComplMessage};
    19 : Result := 'Vous devez renseigner une valeur pour le nombre de jours, de mois ou la base.';
    20 : Result := 'Vous ne pouvez pas saisir de mouvement antérieur à un mois avant la date d''intégration.';
    23 : Result := 'La date de début du mouvement doit être postérieure à la date d''entrée du salarié.';
    26 : Result := 'Vous devez renseigner : le matricule du salarié.';
    27 : Result := 'Vous n''êtes pas autorisé à saisir ce type de mouvement.';
    30 : Result := 'Vous ne pouvez saisir un mouvement à cheval sur plusieurs mois.';
    31 : Result := 'l''heure de début doit précéder l''heure de fin sur une même journée.';
    32 : Result := 'Un mouvement de présence existe déjà sur cette période.'
  else
    Result := '';
  end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 08/08/2007
Modifié le ... : 02/10/2007 GGU
Description .. : Sauvegarde la date du jour traité dans les
Suite ........ : paramètres société.
Mots clefs ... :
*****************************************************************}
Procedure PresenceSauveJourCalcule (DateJour : TDateTime);
Var
  ParamSoc : String;
Begin
  ParamSoc := PresenceRenvoieParamSoc(JOUR);
  If ParamSoc <> '' Then
  Begin
    If DateJour > Date Then DateJour := Date;
    SetParamSoc(ParamSoc, DateToStr(DateJour));
    {$IFDEF EAGLCLIENT}
    AvertirCacheServer('PARAMSOC');
    {$ENDIF}
  End;
End;

//{***********A.G.L.***********************************************
//Auteur  ...... : FLO
//Créé le ...... : 08/08/2007
//Modifié le ... :   /  /
//Description .. : Sauvegarde la date de fin du mois traité dans les
//Suite ........ : paramètres société. Epure en nouvelle année en ne gardant
//Suite ........ : quel le dernier mois de l'année précédente.
//Paramètres ... : Recalcul : Vrai si un recalcul a été lancé
//Mots clefs ... :
//*****************************************************************}
//Procedure PresenceSauveMoisCalcule (Fin : TDateTime; Recalcul : Boolean = False);
//Var
//  A,M,J : Word;
//  Init  : Boolean;
//Begin
//     DecodeDate(Fin, A, M, J);
//
//     // Epuration du champ chaque année (garder le dernier mois de l'année précédente)
//     Init := (M = 1);
//
//     PresenceSauvePeriodeCalculee (MOIS, Fin, Init, Recalcul);
//End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 08/08/2007
Modifié le ... :   /  /
Description .. : Retourne les dates de début et de fin du mois
Suite ........ : en cours.
Mots clefs ... :
*****************************************************************}
Procedure PresenceDonneMoisCalculActuel(var Debut, Fin : TDateTime);
Begin
     PresenceDonnePeriodeCalcul (MOIS, 0, Debut, Fin);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 08/08/2007
Modifié le ... :   /  /
Description .. : Retourne les dates de début et de fin de la semaine
Suite ........ : en cours.
Mots clefs ... :
*****************************************************************}
Procedure PresenceDonneSemaineCalculActuelle(var Debut, Fin : TDateTime);
Begin
     PresenceDonnePeriodeCalcul (SEMAINE, 0, Debut, Fin);
End;

//{***********A.G.L.***********************************************
//Auteur  ...... : FLO
//Créé le ...... : 08/08/2007
//Modifié le ... :   /  /
//Description .. : Méthode permettant la sauvegarde de la
//Suite ........ : date de fin de la période traitée dans les paramètres société.
//Suite ........ : Seules les semaines et mois sont gérés. Le prévisionnel n'est pas enregistré.
//
//Paramètres ... : TypePeriode : JOUR, MOIS
//.............. : Reinit -> True : Réinitialise le paramètre (en gardant le dernier, mois uniquement)
//.............. : Intel  -> True : Si date à sauver est antériere à une date précédente,
//                                  supprime les dates supérieures
//
//Note ......... : Préférer les fonctions SauveMoisCalcule et SauveSemaineCalculee
//.............. : pour les traitements de base.
//Mots clefs ... :
//*****************************************************************}
//Procedure PresenceSauvePeriodeCalculee (TypePeriode : ShortInt; Fin : TDateTime; Reinit : Boolean = False; Intel : Boolean = True);
//Var
//  ListeParam,ParamSoc,NewListe,DateLue : String;
//Begin
//     ParamSoc := PresenceRenvoieParamSoc(TypePeriode);
//
//     Case TypePeriode Of
//          JOUR: // Sauvegarde de la date du jour au max si prévisionnel
//                If Fin > Date Then Fin := Date;
//          MOIS: // Pas de sauvegarde de date de calcul si prévisionnel
//                If Fin > Date Then Exit;
//     End;
//
//     If ParamSoc <> '' Then
//     Begin
//          ListeParam := GetParamSocSecur (ParamSoc, '');
//
//          If (Reinit) And (TypePeriode=MOIS) Then
//               // Récupération dernière date format XX/XX/XXXX
//               ListeParam := Copy(ListeParam, Length(ListeParam)-9, 10)
//          Else
//          Begin
//               // Mode intelligent : compare les dates présentes
//               If (Intel) And (TypePeriode=MOIS) Then
//               Begin
//                    DateLue := ReadTokenSt(ListeParam);
//                    While (DateLue <> '') And (StrToDate(DateLue) < Fin) Do
//                    Begin
//                         If NewListe <> '' Then NewListe := NewListe + ';';
//                         NewListe := NewListe + DateLue;
//                         DateLue := ReadTokenSt(ListeParam);
//                    End;
//                    ListeParam := NewListe;
//               End
//               Else
//               // Fonction de base : Ajout à la fin de la date
//               Begin
//                    If TypePeriode = JOUR Then ListeParam := '';
//               End;
//          End;
//          If ListeParam <> '' Then ListeParam := ListeParam + ';';
//          ListeParam := ListeParam + DateToStr(Fin);
//          SetParamSoc(ParamSoc, ListeParam);
//          {$IFDEF EAGLCLIENT}
//          AvertirCacheServer('PARAMSOC');
//          {$ENDIF}
//     End;
//End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 08/08/2007
Modifié le ... :   /  /
Description .. : Méthode permettant la récupération des dates
Suite ........ : de début et de fin d'une période.
Paramètres ... : TypePeriode : JOUR, SEMAINE, MOIS
.............. : NumPeriode  : 1..12 (pour les mois uniquement)
.............. :               0 : période actuelle
.............. :              -1..-12 (pour les mois uniquement)
Mots clefs ... :
*****************************************************************}
Procedure PresenceDonnePeriodeCalcul (TypePeriode,NumPeriode : ShortInt; var Debut, Fin : TDateTime);
Var DateParam     : TDateTime;
    ListeParam, P : String;
    TempList      : TStringList;
Begin
     Debut := iDate1900;
     Fin   := iDate1900;

     ListeParam := GetParamSocDatesCloture(TypePeriode);  //PT3 GetParamSocSecur(PresenceRenvoieParamSoc(TypePeriode), '');

     Case TypePeriode Of
          JOUR:
               Begin
                    If ListeParam = '' Then
                         // Par défaut : 1er jour du mois actuel
                         DateParam := StartOfTheMonth(Date)
                    Else
                         DateParam := StrToDate(ListeParam) + 1;
                         
                    // La date du jour trouvée, la période s'étend de la dite date jusqu'à aujourd'hui
                    Debut := DateParam;
                    Fin   := Date;
               End;

          SEMAINE:
               Begin
                    If ListeParam = '' Then
                         // Par défaut : Début du mois actuel
                         DateParam := StartOfTheMonth(Date)
                    Else
                         DateParam := StrToDate(ListeParam) + 1;

                    // La date du jour trouvée, la période s'étend sur la semaine correspondante
                    Debut := DebutDeLaSemaine(DateParam);
                    Fin   := FinDeLaSemaine  (DateParam);
               End;

          MOIS:
               Begin
                    If (NumPeriode > 12) Or (NumPeriode < -12) Then Exit;

                    If ListeParam = '' Then
                    Begin
                         // On recherche les calculs de jours déjà effectués
                         ListeParam := GetParamSocDatesCloture(JOUR);  //PT3 GetParamSocSecur(PresenceRenvoieParamSoc(JOUR), '');
                         If ListeParam = '' Then
                              // Par défaut : mois actuel
                              DateParam := StartOfTheMonth(Date)
                         Else
                              // Si on a déjà calculé des jours, on prend le premier jour calculé
                              DateParam := StartOfTheMonth(StrToDate(ListeParam));

                         Debut := DateParam;
                         Fin   := EndOfTheMonth (DateParam);
                    End
                    Else
                    Begin
                         // Création d'une liste des mois traités
                         TempList := TStringList.Create;
                         Try
                              P := ReadTokenSt(ListeParam);
                              While (P <> '') Do
                              Begin
                                   TempList.Add(P);
                                   P := ReadTokenSt(ListeParam);
                              End;

                              // Récupération de l'élèment voulu
                              If (NumPeriode > 0) And (NumPeriode<=TempList.Count) Then
                              Begin
                                   Fin := StrToDate(TempList[NumPeriode-1]);
                                   If (NumPeriode-2 >= 0) Then
                                        Debut := StrToDate(TempList[NumPeriode-2])+1
                                   Else
                                        Debut := IncMonth(Fin,-1)-1;
                              End
                              Else If (NumPeriode = 0) Then
                                   Debut := StrToDate(TempList[TempList.Count-1]) + 1
                              Else If (NumPeriode < 0) And (Abs(NumPeriode)<=TempList.Count) Then
                              Begin
                                   Fin := StrToDate(TempList[TempList.Count+NumPeriode]);
                                   If (Abs(NumPeriode-1) <= TempList.Count) Then
                                        Debut := StrToDate(TempList[TempList.Count+NumPeriode-1])+1
                                   Else
                                        Debut := IncMonth(Fin,-1)-1;
                              End;
                              If Fin = iDate1900 Then Fin := IncMonth(Debut) - 1;
                         Finally
                              TempList.Free;
                         End;
                    End;
               End;
     End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007
Modifié le ... :   /  /
Description .. : Retourne le nom du paramètre société utilisé pour la
Suite ........ : sauvegarde des périodes traitées par le calcul
Mots clefs ... : 
*****************************************************************}
Function PresenceRenvoieParamSoc (TypePeriode : ShortInt) : String;
Begin
     Case TypePeriode Of
          JOUR:
                Result := 'SO_PGPRESJOURSTRAITES';
          SEMAINE:
                Result := 'SO_PGPRESJOURSTRAITES';
          MOIS:
                Result := 'SO_PGPRESMOISTRAITES';
     End;
End;

{ Procedures communes de gestion des plannings de présence }

procedure CreateNewPresenceItem(JourneeType: String; DateCourante : TDateTime; RessourceID, RessourceLibelle: String; TobJourneeType, TobItems : Tob; IsException : Boolean = False; clef : String = '');
var
  TempTobItems, TempTobMvt : Tob;
  WDateDeb, WDateFin : TDateTime;
  procedure RoundHour(var DateTime : TDateTime; RoundSupp : Boolean);
  var
    AYear, AMonth, ADay, AHour, AMinute, ASecond, AMilliSecond : Word;
  begin
    DecodeDateTime(DateTime,AYear, AMonth, ADay, AHour, AMinute, ASecond, AMilliSecond);
    if RoundSupp and (MinuteOf(DateTime) > 0) then
      DateTime := IncHour(DateTime);
    DateTime := RecodeDateTime(DateTime, RecodeLeaveFieldAsIs, RecodeLeaveFieldAsIs, RecodeLeaveFieldAsIs, RecodeLeaveFieldAsIs, 0, 0, 0);
  end;
begin
  TempTobMvt := TobJourneeType.FindFirst(['ETAT'],[JourneeType],False);
  if Assigned(TempTobMvt) then
  begin
    GetDateHeureDebutFin( TempTobMvt.GetDateTime('PJO_HORDEBPLAGE1')
                        , TempTobMvt.GetDateTime('PJO_HORFINPLAGE1')
                        , TempTobMvt.GetDateTime('PJO_HORDEBPLAGE2')
                        , TempTobMvt.GetDateTime('PJO_HORFINPLAGE2')
                        , TempTobMvt.GetString('PJO_JOURJ1PLAGE1')
                        , TempTobMvt.GetString('PJO_JOURJ1PLAGE2')
                        , WDateDeb, WDateFin
                        , DateCourante);
    RoundHour(WDateDeb, False);
    RoundHour(WDateFin, True);
    TempTobItems := Tob.Create('Présence ou Absence',TobItems,-1);
    TempTobItems.AddChampSupValeur(RessourceLibelle,RessourceID);
    TempTobItems.AddChampSupValeur('DATEDEBUT',WDateDeb);
    TempTobItems.AddChampSupValeur('JOUR', DateOf(WDateDeb));
    TempTobItems.AddChampSupValeur('DATEFIN',WDateFin);
    TempTobItems.AddChampSupValeur('ETAT',TempTobMvt.GetString('ETAT'));
    TempTobItems.AddChampSupValeur('ETAT_LIBELLE',TempTobMvt.GetString('ETAT_LIBELLE'));
    TempTobItems.AddChampSupValeur('ETAT_ABREGE',TempTobMvt.GetString('ETAT_ABREGE'));
    TempTobItems.AddChampSupValeur('PJO_HORDEBPLAGE1',TempTobMvt.GetDateTime('PJO_HORDEBPLAGE1'));
    TempTobItems.AddChampSupValeur('PJO_HORFINPLAGE1',TempTobMvt.GetDateTime('PJO_HORFINPLAGE1'));
    TempTobItems.AddChampSupValeur('PJO_HORDEBPLAGE2',TempTobMvt.GetDateTime('PJO_HORDEBPLAGE2'));
    TempTobItems.AddChampSupValeur('PJO_HORFINPLAGE2',TempTobMvt.GetDateTime('PJO_HORFINPLAGE2'));
    TempTobItems.AddChampSupValeur('PJO_JOURJ1PLAGE1',TempTobMvt.GetString('PJO_JOURJ1PLAGE1'));
    TempTobItems.AddChampSupValeur('PJO_JOURJ1PLAGE2',TempTobMvt.GetString('PJO_JOURJ1PLAGE2'));
    TempTobItems.AddChampSupValeur('TYPE_ITEM',TempTobMvt.GetString('TYPE_ITEM'));
    TempTobItems.AddChampSupValeur('CLEF',clef);
  end;
end;

Procedure CompteursARecalculer(DateRecalcule : TDateTime; Salarie : String = 'TOUS');
begin
  ExecuteSQL('Delete PRESENCESALARIE where PYP_SALARIE = "'+Salarie+'" '
            +'and PYP_DATEDEBUTPRES = "'+USDATETIME(DateRecalcule)+'" and PYP_DATEFINPRES = "'+USDATETIME(DateRecalcule)+'" '
            +'and PYP_COMPTEURPRES = "TOUS" and PYP_TYPECALPRES = "001" ');
  ExecuteSQL('Insert into PRESENCESALARIE (PYP_SALARIE, PYP_DATEDEBUTPRES, '
            +'PYP_DATEFINPRES, PYP_COMPTEURPRES, PYP_PERIODICITEPRE, PYP_THEMEPRE, '
            +'PYP_QUANTITEPRES, PYP_ETATPRES, PYP_PGINDICATPRES, PYP_TYPECALPRES, '
            +'PYP_ETABLISSEMENT, PYP_TRAVAILN1, PYP_TRAVAILN2, PYP_TRAVAILN3, '
            +'PYP_TRAVAILN4, PYP_CODESTAT, PYP_CONVENTION, PYP_DADSPROF, PYP_DADSCAT, '
            +'PYP_LIBREPCMB1, PYP_LIBREPCMB2, PYP_LIBREPCMB3, PYP_LIBREPCMB4, '
            +'PYP_DATEDEBUTBUL, PYP_DATEFINBUL, PYP_DATECREATION, PYP_DATEMODIF)'
            +' values ("'+Salarie+'", "'+USDATETIME(DateRecalcule)+'", '
            +'"'+USDATETIME(DateRecalcule)+'", "TOUS", "", "", '
            +'0, "CAL", "ARE", "001", '
            +'"", "", "", "", '                
            +'"", "", "", "", "", '
            +'"", "", "", "", '
            +'"", "", "'+USDATETIME(Date)+'", "'+USDATETIME(Date)+'")');
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/08/2007
Modifié le ... :   /  /
Description .. : Détermination du profil de présence d'un salarié si le type est
Suite ........ : POP (Idem population).
Paramètres ... : LaDate : si 0 alors date du jour
Mots clefs ... :
*****************************************************************}
function GetProfilPresenceFromPop (Salarie : String; LaDate : TDateTime=0) : String;
var GP : TGestionPresence;
Begin
     Result := '';
     GP := TGestionPresence.Create(True,True,False,False,False,0,0,True,Salarie);
     If Assigned(GP) Then
     Begin
          If LaDate = 0 Then LaDate := Date;
          Result := GP.GetProfil(Salarie, LaDate, 'POP');
          FreeAndNil(GP);
     End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 21/08/2007
Modifié le ... :   /  /
Description .. : Mise à jour du profil de présence d'un salarié si le type est
Suite ........ : POP (Idem population).
Paramètres ... : LaDate : si 0 alors date du jour
Mots clefs ... :
*****************************************************************}
procedure UpdateProfilPresenceFromPop (Salarie : String; LaDate : TDateTime=0);
var GP : TGestionPresence;
Begin
     GP := TGestionPresence.Create(True,True,False,False,False,0,0,True,Salarie);
     If Assigned(GP) Then
     Begin
          If LaDate = 0 Then LaDate := Date;
          GP.UpdateProfil(Salarie, LaDate);
          FreeAndNil(GP);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007
Modifié le ... :   /  /
Description .. : Détermine si un modèle ou un cycle est utilisé au sein d'un profil
Mots clefs ... :
*****************************************************************}
Function PresenceCycleIsUsed (TypeAffect, CycleAffect : String) : Boolean;
var
Q : TQUERY;
cycledumodele : string;
Begin
     Result := ExisteSQL('SELECT 1 FROM PROFILPRESENCE WHERE PPQ_TYPEAFFECT="'+TypeAffect+'" AND PPQ_CYCLEAFFECT="'+CycleAffect+'"');
     if Result = false then
     begin      // Recherche si le modèle fait partie d'un cycle
        Q := opensql('SELECT PYD_CYCLE FROM CYCLE WHERE PYD_MODELECYCLE = "'+CycleAffect+'" ',true);
        while not Q.EOF do
        begin    // recherche si ce cycle est affecté
          cycledumodele := Q.findfield('PYD_CYCLE').asstring;
          Result := ExisteSQL('SELECT 1 FROM PROFILPRESENCE WHERE PPQ_TYPEAFFECT= "CYC"  AND PPQ_CYCLEAFFECT="'+CycleDumodele+'"');
          if Result = true then   break;
        Q.next
        end;
     ferme(Q);
     end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007
Modifié le ... :   /  /
Description .. : Détermine si un compteur est utilisé par un profil
Mots clefs ... :
*****************************************************************}
Function PresenceMeterIsUsed (Compteur : String) : Boolean;
Begin
     Result := ExisteSQL('SELECT 1 FROM PROFILCOMPTEURPRES WHERE PPV_COMPTEURPRES="'+Compteur+'"');
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007
Modifié le ... :   /  /
Description .. : Détermine si un profil est rattaché à un salarié
Mots clefs ... :
*****************************************************************}
Function PresenceProfileIsUsed (Profile : String) : Boolean;
Begin
     Result := ExisteSQL('SELECT 1 FROM PROFILPRESSALARIE WHERE PPZ_PROFILPRES="'+Profile+'" AND PPZ_TYPPROFILPRES="PER"');
     If Not Result Then
       Result := ExisteSQL('SELECT 1 FROM PGPARAMETRES WHERE ##PGP_PREDEFINI## AND PGP_MODULECHAMP="'+TYP_POPUL_PRES+'" AND PGP_PGNOMCHAMP="PROFILPRES" AND PGP_PGVALCHAMP="'+Profile+'"');
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007
Modifié le ... :   /  /
Description .. : Détermine si une journée est utilisée au sein d'un modèle de cycle
Mots clefs ... :
*****************************************************************}
Function PresenceJourneyIsUsed (Journey : String) : Boolean;
Begin
     Result := ExisteSQL('SELECT 1 FROM MODELECYCLE WHERE PMO_JOURNEETYPE="'+Journey+'"');
End;


{ TGestionPresence }

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Constructeur de l'objet. Charge toutes les données en rapport
Suite ........ : avec la présence.
Paramètres ... : OnlyPop  : Si oui, ne charge que les données liées aux populations
               : Salarie  : Si présent, ne charge que les données de population du salarié
Mots clefs ... :
*****************************************************************}
constructor TGestionPresence.Create(NeedSalarie : Boolean = True; NeedPopulation : Boolean = True; NeedException : Boolean = True; NeedCycles : Boolean = True; NeedAbsencePresence : Boolean = True; Du : TDateTime = 0; Au : TDateTime = 0; OnlyPop : Boolean=False; Salarie : String = '');
begin
  If Not OnlyPop Then ReloadAssocModele;
  if NeedCycles then
  begin
    ReloadAssocCycle;
    ReloadDureeCycleModele;
    ReloadDureeCycle;
  end;
  ReloadProfilPresence;
  if NeedException then ReloadException;
  if NeedSalarie then
  begin
    ReloadProfilPresenceSalarie;
    if NeedException then ReloadExceptionSal;
    if NeedPopulation then
    begin
      ReloadSalariePopul(Salarie);
      ReloadParametresAssocPopul;
      ReloadParametresPopul;
    end;
  end;
  if NeedAbsencePresence then ReloadAbsencePresenceSalarie(Du, Au, False);
  SavTobModeleAndJournee := tob.Create('Tob de sauvegarde des modèles et journées types calculés pour une date et un cycle', nil, -1);
end;

function TGestionPresence.DateOfException(LaDate: TDateTime; Cycle,
  TypeDeCycle: String): Boolean;
begin
  result := ( GetException(LaDate, Cycle, TypeDeCycle) <> '');
end;

destructor TGestionPresence.Destroy;
begin
  if Assigned(SavTobModeleAndJournee)    then FreeAndNil(SavTobModeleAndJournee);

  if Assigned(TobParametresPopul)        then FreeAndNil(TobParametresPopul);
  if Assigned(TobParametresAssocPopul)   then FreeAndNil(TobParametresAssocPopul);
  if Assigned(TobSalariePopul)           then FreeAndNil(TobSalariePopul);
  if Assigned(TobAssocCycle)             then FreeAndNil(TobAssocCycle);
  if Assigned(TobAssocModele)            then FreeAndNil(TobAssocModele);
  if Assigned(TobDureeCycle)             then FreeAndNil(TobDureeCycle);
  if Assigned(TobDureeCycleModele)       then FreeAndNil(TobDureeCycleModele);
  if Assigned(TobExceptions)             then FreeAndNil(TobExceptions);
  if Assigned(TobExceptionSalaries)      then FreeAndNil(TobExceptionSalaries);
  if Assigned(TobProfilPresenceSalarie)  then FreeAndNil(TobProfilPresenceSalarie);
  if Assigned(TobProfilPresence)         then FreeAndNil(TobProfilPresence);
  if Assigned(TobAbsencePresenceSalarie) then FreeAndNil(TobAbsencePresenceSalarie);

  if Assigned(TobCycle)                  then FreeAndNil(TobCycle);
  if Assigned(TobModeleCycle)            then FreeAndNil(TobModeleCycle);
  inherited;
end;

function TGestionPresence.GetAbsencePresenceSalarie(LaDate: TDateTime; Salarie: String; var TypeMvt, Ressource: String; var Ordre: Integer): String;
var
  TempTob, TobAbsPresSalarie : Tob;
  indexAbsPre : Integer;
begin
  Result := '';
  TobAbsPresSalarie := GetTobAbsencePresenceFiltreSalarie(Salarie);
  if Assigned(TobAbsPresSalarie) then
  begin
    for indexAbsPre := 0 to TobAbsPresSalarie.Detail.Count -1 do
    begin
      TempTob := TobAbsPresSalarie.Detail[indexAbsPre];
      if  (LaDate >= TempTob.GetDateTime('PCN_DATEDEBUTABS'))
      and (LaDate <= TempTob.GetDateTime('PCN_DATEFINABS')) then
      begin
        Ordre := TempTob.GetInteger('PCN_ORDRE');
        Ressource := TempTob.GetString('PCN_RESSOURCE');
        result := TempTob.GetString('PCN_TYPECONGE');
        TypeMvt := TempTob.GetString('PCN_TYPEMVT');
        exit;
      end;
    end;
  end;
//  TempTob := TobAbsencePresenceSalarie.FindFirst(['PCN_SALARIE'], [Salarie], False);
//  while Assigned(TempTob) do
//  begin
//    if  (LaDate >= TempTob.GetDateTime('PCN_DATEDEBUTABS'))
//    and (LaDate <= TempTob.GetDateTime('PCN_DATEFINABS')) then
//    begin
//      Ordre := TempTob.GetInteger('PCN_ORDRE');
//      Ressource := TempTob.GetString('PCN_RESSOURCE');
//      result := TempTob.GetString('PCN_TYPECONGE');
//      TypeMvt := TempTob.GetString('PCN_TYPEMVT');
//      exit;
//    end;
//    TempTob := TobAbsencePresenceSalarie.FindNext(['PCN_SALARIE'], [Salarie], False);
//  end;
end;

function TGestionPresence.GetAbsenceSalarie(LaDate: TDateTime;
  Salarie: String; var Ressource : String; var Ordre : Integer): String;
var
  TempTob, TobAbsPresSalarie : Tob;
begin
  Result := '';
  TobAbsPresSalarie := GetTobAbsencePresenceFiltreSalarie(Salarie);
  if Assigned(TobAbsPresSalarie) then
  begin
    TempTob := TobAbsPresSalarie.FindFirst(['PCN_TYPEMVT'], ['ABS'], False);   // TobAbsencePresenceSalarie   , 'PCN_SALARIE'     , Salarie
    while Assigned(TempTob) do
    begin
      if  (LaDate >= TempTob.GetDateTime('PCN_DATEDEBUTABS'))
      and (LaDate <= TempTob.GetDateTime('PCN_DATEFINABS')) then
      begin
        Ordre := TempTob.GetInteger('PCN_ORDRE');
        Ressource := TempTob.GetString('PCN_RESSOURCE');
        result := TempTob.GetString('PCN_TYPECONGE');
        exit;
      end;
      TempTob := TobAbsPresSalarie.FindNext(['PCN_TYPEMVT'], ['ABS'], False); //  , 'PCN_SALARIE'    , Salarie
    end;
  end;
end;

function TGestionPresence.GetException(LaDate: TDateTime; Cycle,
  TypeDeCycle: String): String;
var
  TobException : Tob;
begin
  result := '';
  TobException := TobExceptions.FindFirst(['PYA_TYPECYCLE', 'PYA_CYCLE', 'PYA_DATEEXCEPTION'], [TypeDeCycle, Cycle, DateOf(LaDate)], False);
  if Assigned(TobException) then result := TobException.GetString('PYA_JOURNEETYPE');
end;

function TGestionPresence.GetExceptionSalarie(LaDate: TDateTime;
  Salarie: String; var CycleOuModele, TypeDeCycle : String): Boolean;
var
  tempTob : Tob;
begin
  result := False;
  CycleOuModele := '';
  TypeDeCycle   := '';
  TempTob := TobExceptionSalaries.FindFirst(['PYE_SALARIE'], [Salarie], False);
  While Assigned(TempTob) do
  begin
    if  (LaDate >= TempTob.GetDateTime('PYE_DATEDEBUT'))
    and (LaDate <= TempTob.GetDateTime('PYE_DATEFIN')) then
    begin
      CycleOuModele := TempTob.GetString('PYE_CYCLEAFFECT');
      TypeDeCycle   := TempTob.GetString('PYE_TYPEAFFECT');
      result := True;
      break;
    end;
    TempTob := TobExceptionSalaries.FindNext(['PYE_SALARIE'], [Salarie], False);
  end;
end;

Function TGestionPresence.GetTobAbsencePresenceSalarie(LaDate : TDateTime; Salarie : String; FiltreTypeMvt : String = '') : Tob;
var
  TempTob, TobAbsPresSalarie : Tob;
begin
  Result := nil;
  TobAbsPresSalarie := GetTobAbsencePresenceFiltreSalarie(Salarie);
  if Assigned(TobAbsPresSalarie) then
  begin
    if FiltreTypeMvt = '' then
    begin
      TempTob := TobAbsPresSalarie.FindFirst([] ,[] ,False);
    end else begin
      TempTob := TobAbsPresSalarie.FindFirst(['PCN_TYPEMVT'] ,[FiltreTypeMvt] ,False);
    end;
  end else TempTob := nil;
  while Assigned(TempTob) do
  begin
    if  (LaDate >= TempTob.GetDateTime('PCN_DATEDEBUTABS'))
    and (LaDate <= TempTob.GetDateTime('PCN_DATEFINABS')) then
    begin
      result := TempTob;
      exit;
    end;
    if FiltreTypeMvt = '' then
      TempTob := TobAbsPresSalarie.FindNext([] ,[] ,False)
    else
      TempTob := TobAbsPresSalarie.FindNext(['PCN_TYPEMVT'] ,[FiltreTypeMvt] ,False)
  end;
end;

Function TGestionPresence.GetTobExceptionSalarie(LaDate: TDateTime;
  Salarie: String): Tob;
var
  tempTob : Tob;
begin
  result := nil;
  TempTob := TobExceptionSalaries.FindFirst(['PYE_SALARIE'], [Salarie], False);
  While Assigned(TempTob) do
  begin
    if  (LaDate >= TempTob.GetDateTime('PYE_DATEDEBUT'))
    and (LaDate <= TempTob.GetDateTime('PYE_DATEFIN')) then
    begin
      result := TempTob;
      break;
    end;
    TempTob := TobExceptionSalaries.FindNext(['PYE_SALARIE'], [Salarie], False);
  end;
end;

function TGestionPresence.GetJourneeType(LaDate : TDateTime; Cycle : String; TypeDeCycle : String = 'CYC'; CheckException : Boolean = True) : String;
var
  Modele, Journee : String;
  DateLancementCycle, DateDebutModele, DateFinModele : TDateTime;
  resteJours : Integer;
begin
  GetModeleAndJourneeInCycle(Modele, Journee, DateLancementCycle, DateDebutModele, DateFinModele, resteJours, LaDate, Cycle, TypeDeCycle, CheckException);
  result := Journee;
end;

procedure TGestionPresence.GetModeleAndJourneeInCycle(var Modele, Journee : String; var DateLancementCycle : TDateTime; var DateDebutModele, DateFinModele : TDateTime; var resteJours : Integer;
  LaDate : TDateTime; Cycle : String; TypeDeCycle : String = 'CYC'; CheckException : Boolean = True) ;
var
  tempTob, tempTobParent, tempTobAssocModele : Tob;
  reste, DiffDates, DureeModele, DureeCycle : Integer;
  stJourneeTypeException : String;
begin
  { Pour optimiser les traitements, les résultats sont stockés dans une tob
    et ressortis à la demande. }
  tempTob := SavTobModeleAndJournee.FindFirst(['TYPEDECYCLE'], [TypeDeCycle], False);
  if Assigned(tempTob) then
  begin
    tempTob := tempTob.FindFirst(['CYCLE'], [Cycle], False);
    if Assigned(tempTob) then
    begin
      tempTob := tempTob.FindFirst(['CHECKEXCEPT'], [CheckException], False);
      if Assigned(tempTob) then
      begin
        tempTob := tempTob.FindFirst(['DATE'], [LaDate], False);
        if Assigned(tempTob) then
        begin
          Modele             := tempTob.GetString('MODELE');
          Journee            := tempTob.GetString('JOURNEE');
          DateLancementCycle := tempTob.GetDateTime('LANCEMENTCYCLE');
          DateDebutModele    := tempTob.GetDateTime('DEBUTMODELE');
          DateFinModele      := tempTob.GetDateTime('FINMODELE');
          resteJours         := tempTob.GetInteger('RESTEJOUR');
          exit;
        end;
      end;
    end;
  end;

  Modele := '';
  Journee := '';
  DateLancementCycle := iDate1900;
  DateDebutModele := iDate1900;
  DateFinModele := iDate2099;
  resteJours := 0;
  if TypeDeCycle = 'MOD' then //Affichage des modèles de cycle  (Type défini dans la tablette PGTYPEAFFECTSALARIE)
  begin
    Modele := Cycle;
    { si Modèle, on cherche la durée du modèle }
    tempTob := TobAssocModele.FindFirst(['PMY_MODELECYCLE'],[Modele],False); { on prends n'importe quel ligne qui correspond au modele }
    if Assigned(tempTob) and (tempTob.GetString('PMY_DATEVALIDITE') <> '') and (tempTob.GetString('PMY_NBJOUR') <> '') then
    begin
      DateLancementCycle := tempTob.GetDateTime('PMY_DATEVALIDITE');
      DiffDates := Trunc(LaDate) - Trunc(DateLancementCycle) +1;
      if DiffDates < 1 then exit; { le modèle n'est pas valide avant sa date de validité }
      { on calcul le reste }
      DureeModele := tempTob.GetInteger('PMY_NBJOUR');
      reste := DiffDates MOD DureeModele;
        if reste = 0 then reste := DureeModele;
      { Recherche du jour concerné dans le modèle }
      tempTob := TobAssocModele.FindFirst(['PMY_MODELECYCLE','PMO_ORDREJOUR'], [Modele, reste], False);

      DateDebutModele := DateLancementCycle + DureeModele * Trunc((Trunc(LaDate) - Trunc(DateLancementCycle))/DureeModele);
      DateFinModele := DateDebutModele + DureeModele -1;

    end;
  end else begin
    { si cycle, on cherche la durée du cycle }
    tempTob := TobDureeCycle.FindFirst(['PYC_CYCLE'],[Cycle],False);
    if Assigned(tempTob) and (tempTob.GetString('DUREE') <> '') then
    begin
      { on calcul le reste }
      DateLancementCycle := tempTob.GetDateTime('DATE');
      DiffDates := Trunc(LaDate) - Trunc(DateLancementCycle) +1;
      if DiffDates < 1 then exit; { le cycle n'est pas valide avant sa date de validité }
      DureeCycle := tempTob.GetInteger('DUREE');
      reste := DiffDates MOD DureeCycle;
      DateDebutModele := DateLancementCycle + DureeCycle * Trunc((Trunc(LaDate) - Trunc(DateLancementCycle))/DureeCycle);
      { on en déduit le modèle concerné }
      tempTob := TobDureeCycleModele.FindFirst(['PYC_CYCLE'], [Cycle], False);
      if Assigned(tempTob) and (tempTob.GetString('DUREE') <> '') then
        DureeModele := tempTob.GetInteger('DUREE')
      else
        DureeModele := 0;
      DateFinModele := DateDebutModele + DureeModele -1;
      While (reste > DureeModele) do
      begin
        reste := reste - DureeModele;
        tempTob := TobDureeCycleModele.FindNext(['PYC_CYCLE'], [Cycle], False);
        if Assigned(tempTob) and (tempTob.GetString('DUREE') <> '') then
          DureeModele := tempTob.GetInteger('DUREE')
        else
          DureeModele := 0;
        DateDebutModele := DateFinModele + 1;
        DateFinModele := DateDebutModele + DureeModele -1;
      end;
      if Assigned(tempTob) then
      begin
        Modele := tempTob.GetString('PMY_MODELECYCLE');
        tempTobAssocModele := TobAssocModele.findFirst(['PMY_MODELECYCLE'],[Modele],False);
        if Assigned(tempTobAssocModele) then
          DureeModele := tempTobAssocModele.GetInteger('PMY_NBJOUR');
        if DureeModele <> 0 then
        begin
          reste := reste MOD DureeModele;
          if reste = 0 then reste := DureeModele;
          { Recherche du jour concerné dans le modèle }
          tempTob := TobAssocCycle.FindFirst(['PYC_CYCLE','PMY_MODELECYCLE','PMO_ORDREJOUR'], [Cycle, Modele, reste], False);
        end;
      end;
    end;
  end;
  if Assigned(tempTob) then
    Journee := tempTob.GetString('PMO_JOURNEETYPE');
  if CheckException then
  begin
    stJourneeTypeException := GetException(LaDate, Cycle, TypeDeCycle);
    if stJourneeTypeException <> '' then
      Journee := stJourneeTypeException;
  end;
  { Sauvegarde dans la tob }
  tempTobParent := SavTobModeleAndJournee;
  tempTob := tempTobParent.FindFirst(['TYPEDECYCLE'], [TypeDeCycle], False);
  if not Assigned(tempTob) then
  begin
    tempTob := TOB.Create('Liste des types de cycles', tempTobParent, -1);
    tempTob.AddChampSupValeur('TYPEDECYCLE',TypeDeCycle);
  end;
  tempTobParent := tempTob;
  tempTob := tempTob.FindFirst(['CYCLE'], [Cycle], False);
  if not Assigned(tempTob) then
  begin
    tempTob := TOB.Create('Liste des cycles', tempTobParent, -1);
    tempTob.AddChampSupValeur('CYCLE',Cycle);
  end;
  tempTobParent := tempTob;
  tempTob := tempTob.FindFirst(['CHECKEXCEPT'], [CheckException], False);
  if not Assigned(tempTob) then
  begin
    tempTob := TOB.Create('Exceptions vérifiés ou non', tempTobParent, -1);
    tempTob.AddChampSupValeur('CHECKEXCEPT',CheckException);
  end;
  tempTobParent := tempTob;
  tempTob := tempTob.FindFirst(['DATE'], [LaDate], False);
  if not Assigned(tempTob) then
  begin
    tempTob := TOB.Create('Dates et valeurs', tempTobParent, -1);
    tempTob.AddChampSupValeur('DATE',LaDate);
  end;
  tempTob.AddChampSupValeur('MODELE',Modele);
  tempTob.AddChampSupValeur('JOURNEE',Journee);
  tempTob.AddChampSupValeur('LANCEMENTCYCLE',DateLancementCycle);
  tempTob.AddChampSupValeur('DEBUTMODELE',DateDebutModele);
  tempTob.AddChampSupValeur('FINMODELE',DateFinModele);
  tempTob.AddChampSupValeur('RESTEJOUR',resteJours);
end;

function TGestionPresence.GetJourneeTypeSalarie(LaDate: TDateTime;
  Salarie: String; var CycleOuModele, TypeDeCycle : String; var isExceptionSalarie, isExceptionCycle : Boolean; CheckException: Boolean = True): String;
var
  ProfilPres : String;
  TempTob : Tob;
begin
  isExceptionSalarie := False;
  isExceptionCycle := False;
  result := '';
  { Recherche des exceptions de présence du salarié }
  GetExceptionSalarie( LaDate, Salarie, CycleOuModele, TypeDeCycle);
  isExceptionSalarie := True;
  if TypeDeCycle = 'JOU' then
  begin
    result := CycleOuModele;
    exit;
  end;
  if (result = '') and (CycleOuModele = '') then
  begin
    { Recherche du profil de présence du salarié }
    isExceptionSalarie := False;
    ProfilPres := GetProfil(Salarie, LaDate);

    { lecture du profil de présence du salarié valide à la date de recherche }
    TempTob := GetLastValideFromTob(LaDate,'PPQ_DATEVALIDITE', TobProfilPresence, ['PPQ_PROFILPRES'], [ProfilPres]);
    if Assigned(TempTob) then
    begin
      CycleOuModele := TempTob.GetString('PPQ_CYCLEAFFECT');
      TypeDeCycle   := TempTob.GetString('PPQ_TYPEAFFECT');
    end;
  end;
  { Recherche de la journée type en fonction du cycle ou modèle }
  if (result = '') and (CycleOuModele <> '')then
  begin
    Result := GetJourneeType(LaDate, CycleOuModele, TypeDeCycle, CheckException);
    isExceptionCycle := DateOfException(LaDate, CycleOuModele, TypeDeCycle);
  end;
end;

{function TGestionPresence.GetExceptionsPresenceSalarie(LaDate: TDateTime;
  Salarie: String; var CycleOuModele, TypeDeCycle : String; var isExceptionSalarie, isExceptionCycle : Boolean; CheckException: Boolean = True): String;
var
  TypeProfilPres, ProfilPres : String;
  TempTob : Tob;
  SalariePop, parametrepop : String;
begin
  isExceptionSalarie := False;
  isExceptionCycle := False;
  result := '';
  GetExceptionSalarie( LaDate, Salarie, CycleOuModele, TypeDeCycle);
  isExceptionSalarie := True;
  if TypeDeCycle = 'JOU' then
  begin
    result := CycleOuModele;
    exit;
  end;
  if (result = '') and (CycleOuModele <> '')then
  begin
    Result := GetJourneeType(LaDate, CycleOuModele, TypeDeCycle, CheckException);
    isExceptionCycle := DateOfException(LaDate, CycleOuModele, TypeDeCycle);
  end;
end;  }


function TGestionPresence.GetProfilsPresenceSalarie(Salarie: String): tob;
begin
  result := GetAllValidesFromTobValidite(TobProfilPresenceSalarie, ['PPZ_SALARIE'], [Salarie], 'PPZ_DATEVALIDITE', 'DATEDEBUT', iDate1900, 'DATEFIN', iDate2099 )
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 21/08/2007
Modifié le ... :   /  /
Description .. : Récupération du profil associé à un salarié.
Suite ........ : Si TypeProfilPres non paramétré, on recherche l'information
Suite ........ : dans les données présentes en table. S'il est à POP,  on
Suite ........ : calcule le profil à partir de la population d'appartenance.
Mots clefs ... :
*****************************************************************}
function TGestionPresence.GetProfil (Salarie : String; LaDate : TDateTime; TypeProfilPres : String='' ) : String;
var TempTob : TOB;
    SalariePop, ParametrePop : String;

    function recupProfilPop : String;
    Begin
          Result := '';
          { Recherche de la population du salarié }
          TempTob := TobSalariePopul.FindFirst(['PNA_SALARIE'], [Salarie], False);
          If Assigned(TempTob) then
          begin
               SalariePop := TempTob.GetString('PNA_POPULATION');
               { Recherche du parametre de la population }
               TempTob := GetLastValideFromTob(LaDate,'PGO_DATEVALIDITE', TobParametresAssocPopul, ['PGO_CODEASSOCIE'], [SalariePop]);
               If Assigned(TempTob) Then
               Begin
                    ParametrePop := TempTob.GetString('PGO_PGPARAMETRE');
                    { Recherche du profil de présence du parametre }
                    TempTob := TobParametresPopul.FindFirst(['PGP_PGPARAMETRE', 'PGP_PGNOMCHAMP'], [ParametrePop, 'PROFILPRES'], False);
                    If Assigned(TempTob) Then Result := TempTob.GetString('PGP_PGVALCHAMP');
               End;
          End;
    End;
Begin
     { Cas 1 : Données présentes dans les TOBs }
     If TypeProfilPres = '' Then
     Begin
          { Recherche du profil de présence du salarié }
          TempTob := GetLastValideFromTob(LaDate,'PPZ_DATEVALIDITE', TobProfilPresenceSalarie, ['PPZ_SALARIE'], [Salarie]);
          If Assigned(TempTob) Then
          Begin
               TypeProfilPres := TempTob.GetString('PPZ_TYPPROFILPRES');
               { Profil personnalisé : Rien à faire }
               If TypeProfilPres = 'PER' Then
                    Result := TempTob.GetString('PPZ_PROFILPRES')
               { Profil idem Pop }
               Else If TypeProfilPres = 'POP' Then
                    Result := recupProfilPop;
          End;
     End
     { Cas 2 : Détermination du profil associé au salarié par la population }
     Else If TypeProfilPres = 'POP' Then
          Result := recupProfilPop;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 22/08/2007
Modifié le ... :   /  /
Description .. : Actualise le profil de présence rattaché au salarié dans le 
Suite ........ : cas d'un idem pop
Mots clefs ... : 
*****************************************************************}
Function TGestionPresence.UpdateProfil (Salarie : String; LaDate : TDateTime) : Boolean;
var TobProfils,T : TOB;
    Profil : String;
Begin
     Profil := '';
     // Récupération des profils de présence du salarié
     TobProfils := GetProfilsPresenceSalarie(Salarie);
     If Assigned(TobProfils) And (TobProfils.Detail.Count > 0) Then
     Begin
          // Récupération du profil le plus récent (inutile de mettre à jour d'anciens profils)
          TobProfils.Detail.Sort('PPZ_DATEVALIDITE');
          T := TobProfils.Detail[TobProfils.Detail.Count-1];

          // Mise à jour du profil en cas d'idem pop
          If (T.GetValue('PPZ_TYPPROFILPRES') = 'POP') And (T.GetValue('PPZ_DATEVALIDITE') <= LaDate) Then
          Begin
               Profil := GetProfil(Salarie, LaDate);
               If Profil <> '' Then
               Begin
                    T.PutValue('PPZ_PROFILPRES', Profil);
                    //GGU : La tob T est virtuelle, on ne peut pas utilisé UpdateDB :  // T.UpdateDB();
                    ExecuteSQL(Format('Update PROFILPRESSALARIE set PPZ_PROFILPRES = "%s" where PPZ_SALARIE = "%s" and PPZ_DATEVALIDITE = "%s"', [Profil, Salarie, USDATETIME(T.GetDateTime('PPZ_DATEVALIDITE'))]));
               End;
          End
          Else
               // Pas de profil idem pop actif à cette date : pas d'erreur
               Profil := 'xxx';
     End
     Else
          // Pas de profil paramétré : pas d'erreur
          Profil := 'xxx';
     If Assigned(TobProfils) Then FreeAndNil(TobProfils);
     Result := (Profil<>'');
End;

function TGestionPresence.GetTypesEtProfilsPresenceSalarie(Salarie: String; TobProfilsPresenceSalarie: tob): tob;
var
  tobRythmes, tobRythme, tobProfil, TempTob, TobParamsPopul, TobParametre : Tob;
  indexProfil, indexParamPopul : Integer;
  TypeProfilPres, ProfilPres, SalariePop, parametrepop : String;
begin
  tobRythmes := TOB.Create('Profils de présence', nil, -1);
  result := tobRythmes;
  if Assigned(TobProfilsPresenceSalarie) then
  begin
    for indexProfil := 0 to TobProfilsPresenceSalarie.FillesCount(0) -1 do
    begin
      tobProfil := TobProfilsPresenceSalarie.Detail[indexProfil];
      if Assigned(tobProfil) then
      begin
        TypeProfilPres := tobProfil.GetString('PPZ_TYPPROFILPRES');
        if TypeProfilPres = 'PER' then
        begin
          ProfilPres := tobProfil.GetString('PPZ_PROFILPRES');
          tobRythme := tob.create('Rythme de travail', tobRythmes, -1);
          tobRythme.AddChampSupValeur('PPZ_TYPPROFILPRES', TypeProfilPres);
          tobRythme.AddChampSupValeur('PPZ_PROFILPRES', ProfilPres);
          tobRythme.AddChampSupValeur('DATEDEBUT', tobProfil.GetDateTime('PPZ_DATEVALIDITE'));
          tobRythme.AddChampSupValeur('DATEFIN', tobProfil.GetDateTime('DATEFIN'));
        end else if TypeProfilPres = 'POP' then
        begin
          { Recherche de la population du salarié }
          TempTob := TobSalariePopul.FindFirst(['PNA_SALARIE'], [Salarie], False);
          if Assigned(TempTob) then
          begin
            SalariePop := TempTob.GetString('PNA_POPULATION');
            { Recherche des parametres de la population }
            TobParamsPopul := GetAllValidesFromTobValidite(TobParametresAssocPopul, ['PGO_CODEASSOCIE'], [SalariePop], 'PGO_DATEVALIDITE', 'DATEDEBUT', tobProfil.GetDateTime('PPZ_DATEVALIDITE'), 'DATEFIN', tobProfil.GetDateTime('DATEFIN'));
            for indexParamPopul := 0 to TobParamsPopul.FillesCount(0) -1 do
            begin
              parametrepop := TobParamsPopul.Detail[indexParamPopul].GetString('PGO_PGPARAMETRE');
              { Recherche du profil de présence du parametre }
              TobParametre := TobParametresPopul.FindFirst(['PGP_PGPARAMETRE', 'PGP_PGNOMCHAMP'], [parametrepop, 'PROFILPRES'], False);
              if Assigned(TobParametre) then
              begin
                ProfilPres := TobParametre.GetString('PGP_PGVALCHAMP');
                tobRythme := tob.create('Profil de présence', tobRythmes, -1);
                tobRythme.AddChampSupValeur('PPZ_TYPPROFILPRES', TypeProfilPres);
                tobRythme.AddChampSupValeur('PPZ_PROFILPRES', ProfilPres);
                tobRythme.AddChampSupValeur('DATEDEBUT', TobParamsPopul.Detail[indexParamPopul].GetDateTime('DATEDEBUT'));
                tobRythme.AddChampSupValeur('DATEFIN', TobParamsPopul.Detail[indexParamPopul].GetDateTime('DATEFIN'));
              end;
            end;
            FreeAndNil(TobParamsPopul);
          end;
        end;
      end;
    end;
  end;
end;

function TGestionPresence.GetRythmesSalarie(TobTypesEtProfilsPresenceSalarie : tob): tob; //Salarie : String; 
var
  indexProfil, indexRythme : Integer;
  TobProfil, TobRythme : Tob;
begin
  result := tob.Create('Rythme de travail',nil,-1);
  for indexProfil := 0 to TobTypesEtProfilsPresenceSalarie.FillesCount(0)-1 do
  begin
    TobProfil := TobTypesEtProfilsPresenceSalarie.Detail[indexProfil];
    { lecture du profil de présence du salarié valide à la date de recherche }
    if Assigned(TobProfil) then
    begin
      TobRythme := GetAllValidesFromTobValidite(TobProfilPresence, ['PPQ_PROFILPRES'], [TobProfil.GetString('PPZ_PROFILPRES')], 'PPQ_DATEVALIDITE',
                                                        'DATEDEBUT', TobProfil.GetDateTime('DATEDEBUT'), 'DATEFIN', TobProfil.GetDateTime('DATEFIN'));
      for indexRythme := TobRythme.FillesCount(0)-1 downto 0 do
      begin
        TobRythme.detail[indexRythme].ChangeParent(result, -1);
      end;
      FreeAndNil(TobRythme);
    end;
  end;
end;

function TGestionPresence.GetCyclesSalarie(TobRythmesSalarie : tob): tob; //Salarie : String;
var
  indexProfil, indexCycle : Integer;
  TobProfil : Tob;
  TempTobCycle, temptob : Tob;
begin
  result := tob.Create('Cycle de présence',nil,-1);
  for indexProfil := 0 to TobRythmesSalarie.FillesCount(0)-1 do
  begin
    TobProfil := TobRythmesSalarie.Detail[indexProfil];
    { recherche des cycles ou modèles valides à la date de recherche }
    if Assigned(TobProfil) then
    begin
      if TobProfil.GetString('PPQ_TYPEAFFECT') = 'CYC' then
      begin
        TempTobCycle := GetAllValidesFromTobValidite(TobCycle, ['PYC_CYCLE'], [TobProfil.GetString('PPQ_CYCLEAFFECT')], 'PYC_DATEVALIDITE',
                                                          'DATEDEBUT', TobProfil.GetDateTime('DATEDEBUT'), 'DATEFIN', TobProfil.GetDateTime('DATEFIN'));
        for indexCycle := TempTobCycle.FillesCount(0)-1 downto 0 do
        begin
          temptob := TempTobCycle.detail[indexCycle];
          temptob.AddChampSupValeur('DATEVALIDITE', temptob.GetString('PYC_DATEVALIDITE'));
          temptob.AddChampSupValeur('CYCLE', temptob.GetString('PYC_CYCLE'));
          temptob.AddChampSupValeur('TYPE', 'CYC');
          temptob.AddChampSupValeur('LIBELLE', temptob.GetString('PYC_LIBELLE'));
          temptob.ChangeParent(result, -1);
        end;
      end else if TobProfil.GetString('PPQ_TYPEAFFECT') = 'MOD' then begin
        TempTobCycle := GetAllValidesFromTobValidite(TobModeleCycle, ['PMY_MODELECYCLE'], [TobProfil.GetString('PPQ_CYCLEAFFECT')], 'PMY_DATEVALIDITE',
                                                          'DATEDEBUT', TobProfil.GetDateTime('DATEDEBUT'), 'DATEFIN', TobProfil.GetDateTime('DATEFIN'));
        for indexCycle := TempTobCycle.FillesCount(0)-1 downto 0 do
        begin
          temptob := TempTobCycle.detail[indexCycle];
          temptob.AddChampSupValeur('DATEVALIDITE', temptob.GetString('PMY_DATEVALIDITE'));
          temptob.AddChampSupValeur('CYCLE', temptob.GetString('PMY_MODELECYCLE'));
          temptob.AddChampSupValeur('TYPE', 'MOD');
          temptob.AddChampSupValeur('LIBELLE', temptob.GetString('PMY_LIBELLE'));
          temptob.ChangeParent(result, -1);
        end;
      end;
      FreeAndNil(TempTobCycle);
    end;
  end;
end;

//function TGestionPresence.GetJourneeTypeFromRythmeSalarie(Salarie : String; TobRythmeSalarie : tob): tob;
//begin
//  { Recherche de la journée type en fonction du cycle ou modèle }
//  if (result = '') and (CycleOuModele <> '')then
//  begin
//    Result := GetJourneeType(LaDate, CycleOuModele, TypeDeCycle, CheckException);
//    isExceptionCycle := DateOfException(LaDate, CycleOuModele, TypeDeCycle);
//  end;
//end;

function TGestionPresence.GetPresenceSalarie(LaDate: TDateTime; Salarie: String; var Ressource : String; var Ordre : Integer): String;
var
  TempTob, TobAbsPresSalarie : Tob;
begin
  Result := '';
  TobAbsPresSalarie := GetTobAbsencePresenceFiltreSalarie(Salarie);
  if Assigned(TobAbsPresSalarie) then
  begin
    TempTob := TobAbsPresSalarie.FindFirst( ['PCN_TYPEMVT'], ['ABS'] , False);
    while Assigned(TempTob) do
    begin
      if  (LaDate >= TempTob.GetDateTime('PCN_DATEDEBUTABS'))
      and (LaDate <= TempTob.GetDateTime('PCN_DATEFINABS')) then
      begin
        Ordre := TempTob.GetInteger('PCN_ORDRE');
        Ressource := TempTob.GetString('PCN_RESSOURCE');
        result := TempTob.GetString('PCN_TYPECONGE');
        exit;
      end;
      TempTob := TobAbsPresSalarie.FindNext( ['PCN_TYPEMVT'], ['ABS'], False);
    end;
  end;
end;

procedure TGestionPresence.ReloadAbsencePresenceSalarie(Du, Au : TDateTime; Append : Boolean = True; FiltreSalarie : String = '');
var
  stSQL : String;
  DateDu, DateAu : TDateTime;
  tempTob, NewAbsencesTob, TempTobSal : Tob;
  indextob : Integer;
begin
  DateDu := DateOf(Du);
  DateAu := DateOf(Au);
  if Append then
  begin
    if (DateDu > InMemoryEndDate) or (DateAu < DateDu) or (InMemoryBeginDate = 0) or (InMemoryEndDate = 0) then //Pour des dates complètement séparées, on prends les nouvelles dates
    begin
      InMemoryBeginDate := DateDu;
      InMemoryEndDate := DateAu;
    end else begin //Sinon, on étends les dates existantes
      if (DateAu >= InMemoryBeginDate) and (DateAu <= InMemoryEndDate) then
        InMemoryBeginDate := Min(InMemoryBeginDate, DateDu);
      if (DateDu >= InMemoryBeginDate) and (DateDu <= InMemoryEndDate) then
        InMemoryEndDate := Max(InMemoryEndDate, DateAu);
      if (DateDu < InMemoryBeginDate) and (DateAu > InMemoryEndDate) then
      begin
        InMemoryBeginDate := DateDu;
        InMemoryEndDate := DateAu;
      end;
    end;
  end else begin
    InMemoryBeginDate := DateDu;
    InMemoryEndDate := DateAu;
  end;

  { On éclate la tob des absences-présences par salarié }
  NewAbsencesTob := TOB.Create('Liste des absences et des présences des salariés', nil, -1);
  stSQL := 'SELECT PCN_TYPEMVT, PCN_RESSOURCE, PCN_SALARIE, '
         + 'PCN_ORDRE, PCN_TYPECONGE, PCN_LIBELLE, PCN_DATEDEBUTABS, '
         + 'PCN_HDEB, PCN_DATEFINABS, PCN_HFIN, PCN_HEURES, PCN_JOURS, PCN_DEBUTDJ, PCN_FINDJ, '
         + 'PCN_ETABLISSEMENT, PCN_TRAVAILN1, PCN_TRAVAILN2, PCN_TRAVAILN3, PCN_TRAVAILN4, PCN_CODESTAT, '
         + 'PCN_CONFIDENTIEL, PCN_NBHEURESNUIT, PMA_MOTIFABSENCE, PMA_JOURHEURE '
         + 'FROM ABSENCESALARIE LEFT JOIN MOTIFABSENCE ON PCN_TYPECONGE = PMA_MOTIFABSENCE ';
  if (DateDu > 10) and (DateAu > 10) then
  begin
    stSQL := stSQL + ' WHERE ((PCN_DATEDEBUTABS BETWEEN "'+USDATETIME(DateDu)+'" AND "'+USDATETIME(DateAu)+'") ';
    stSQL := stSQL + ' or    (PCN_DATEFINABS BETWEEN "'+USDATETIME(DateDu)+'" AND "'+USDATETIME(DateAu)+'") ';
    stSQL := stSQL + ' or    ((PCN_DATEDEBUTABS <= "'+USDATETIME(DateDu)+'") AND (PCN_DATEFINABS >= "'+USDATETIME(DateAu)+'")) )';
  end;
  if FiltreSalarie <> '' then
  begin
    stSQL := stSQL + ' AND (PCN_SALARIE = "'+FiltreSalarie+'")';
  end else
    AddFiltreSalarie('PCN', stSQL);
  NewAbsencesTob.LoadDetailFromSQL(stSQL);
  if append then
  begin
    //On vérifie que ces absences n'ont pas déjà été intégrées
    for indextob := NewAbsencesTob.FillesCount(0)-1 downto 0 do
    begin
      tempTob := NewAbsencesTob.Detail[indextob];
      TempTobSal := GetTobAbsencePresenceFiltreSalarie(tempTob.GetString('PCN_SALARIE'));
      if Assigned(TempTobSal) then
        TempTobSal := TempTobSal.FindFirst(['PCN_TYPEMVT' , 'PCN_ORDRE', 'PCN_RESSOURCE']
          ,[tempTob.GetString('PCN_TYPEMVT'), tempTob.GetString('PCN_ORDRE'), tempTob.GetString('PCN_RESSOURCE')], False);
      if Assigned(TempTobSal) then
      begin
        tempTob.ChangeParent(nil,-1);
        FreeAndNil(tempTob);
      end;
    end;
  end else begin
    if Assigned(TobAbsencePresenceSalarie) then FreeAndNil(TobAbsencePresenceSalarie);
  end;
  EclateTob(NewAbsencesTob, TobAbsencePresenceSalarie, 'PCN_SALARIE', False);
  FreeAndNil(NewAbsencesTob);
end;

procedure TGestionPresence.ReloadAssocCycle;
begin
  TobAssocCycle := TOB.Create('Association Cycle -> Modele -> Journée', nil, -1);
  TobAssocCycle.LoadDetailFromSQL( 'select PYC_CYCLE,PYC_LIBELLE, PYC_DATEVALIDITE, PYC_NBCYCLE, '
                                 + ' PYD_ORDRECYCLE, PYD_MODELECYCLE, PMY_MODELECYCLE, PMO_ORDREJOUR, PMO_JOURNEETYPE '
                                 + ' from cycleentete '
                                 + ' left outer join cycle '
                                 + ' on PYC_CYCLE = PYD_CYCLE '
                                 + ' left outer join modelecycleent '
                                 + ' on PYD_MODELECYCLE = PMY_MODELECYCLE '
                                 + ' left outer join modelecycle '
                                 + ' on PYD_MODELECYCLE = PMO_MODELECYCLE');
  TobCycle := TOB.Create('Cycles', nil, -1);
  TobCycle.LoadDetailFromSQL('SELECT PYC_CYCLE, PYC_LIBELLE, PYC_DATEVALIDITE, PYC_TYPECYCLE, PYC_ETATCYCLE, PYC_NBCYCLE, "PMY_NBJOUR", PYC_COLORCYCLE FROM CYCLEENTETE');
  TobModeleCycle := TOB.Create('Modèle de Cycle', nil, -1);
  TobModeleCycle.LoadDetailFromSQL('SELECT PMY_MODELECYCLE, PMY_LIBELLE, PMY_DATEVALIDITE, "PYC_TYPECYCLE", "PYC_ETATCYCLE", "PYC_NBCYCLE", PMY_NBJOUR, PMY_COLORMODCYCLE FROM MODELECYCLEENT');

end;

procedure TGestionPresence.ReloadAssocModele;
begin
  TobAssocModele := TOB.Create('Association Modele -> Journée', nil, -1);
  TobAssocModele.LoadDetailFromSQL( 'select PMY_MODELECYCLE, PMY_DATEVALIDITE, PMY_NBJOUR, PMY_LIBELLE, PMO_ORDREJOUR, PMO_JOURNEETYPE '
                                  + ' from modelecycleent '
                                  + ' left outer join modelecycle '
                                  + ' on PMY_MODELECYCLE = PMO_MODELECYCLE '
                                  + 'Order by PMY_MODELECYCLE, PMO_ORDREJOUR');
end;

procedure TGestionPresence.ReloadDureeCycle;
var
  index, Duree : Integer;
  TempTob :Tob;
  Cycle : String;
begin
  { Calcul des durées total des cyles }
  TobDureeCycle := TOB.Create('Durée et date de référence de chaque cycle', nil, -1);
  For index := 0 to TobDureeCycleModele.FillesCount(0) -1 do
  begin
    if TobDureeCycleModele.Detail[index].GetString('DUREE') <> '' then
    begin
      Cycle := TobDureeCycleModele.Detail[index].GetString('PYC_CYCLE');
      Duree := TobDureeCycleModele.Detail[index].GetInteger('DUREE');
      TempTob := TobDureeCycle.FindFirst(['PYC_CYCLE'],[Cycle],false);
      if Assigned(TempTob) then
      begin
        TempTob.SetInteger('DUREE',TempTob.GetInteger('DUREE')+Duree);
      end else begin
        TempTob := Tob.Create('durée cycle',TobDureeCycle,-1);
        TempTob.AddChampSupValeur('PYC_CYCLE',Cycle);
        TempTob.AddChampSupValeur('DUREE',Duree);
        TempTob.AddChampSupValeur('DATE',TobDureeCycleModele.Detail[index].GetDateTime('PYC_DATEVALIDITE'));
      end;
    end;
  end;
end;

procedure TGestionPresence.ReloadDureeCycleModele;
begin
  TobDureeCycleModele := TOB.Create('Durée de chaque couple cycle/modèle', nil, -1);
  TobDureeCycleModele.LoadDetailFromSQL( 'select pyc_cycle, PYC_DATEVALIDITE, pyd_ordrecycle, PMY_MODELECYCLE, pyd_nbmodcycle * sum(PMY_NBJOUR) as DUREE '
                                       + ' from cycleentete '
                                       + ' left outer join cycle '
                                       + ' on PYC_CYCLE = PYD_CYCLE '
                                       + ' left outer join modelecycleent '
                                       + ' on PYD_MODELECYCLE = PMY_MODELECYCLE '
                                       + ' group by pyc_cycle, PYC_DATEVALIDITE, PMY_MODELECYCLE, pyd_nbmodcycle, pyd_ordrecycle '
                                       + ' order by pyc_cycle, pyd_ordrecycle');
end;

procedure TGestionPresence.ReloadException;
begin
  if not Assigned(TobExceptions) then //FreeAndNil(TobExceptions);
    TobExceptions := Tob.Create('Exceptions ', nil, -1);
  TobExceptions.LoadDetailDBFromSQL('EXCEPTCYCLE','SELECT PYA_TYPECYCLE, PYA_CYCLE, PYA_LIBELLECYCLE, PYA_DATEEXCEPTION, PYA_JOURNEETYPE FROM EXCEPTCYCLE'); //, PYA_DATECREATION, PYA_DATEMODIF, PYA_CREATEUR, PYA_UTILISATEUR
end;

procedure TGestionPresence.ReloadExceptionSal;
var
  stSQL : String;
begin
  if not Assigned(TobExceptionSalaries) then
    TobExceptionSalaries := TOB.Create('Exceptions de présence du salarié', nil, -1);
  stSQL := 'SELECT PYE_SALARIE, PYE_DATEDEBUT, PYE_DATEFIN, PYE_TYPEAFFECT, PYE_CYCLEAFFECT '
         + ' FROM EXCEPTPRESENCESAL ';
  AddFiltreSalarie('PYE', stSQL);
  TobExceptionSalaries.LoadDetailFromSQL(stSQL);
end;

procedure TGestionPresence.ReloadParametresAssocPopul;
begin
  { Chargement de la table d'association Population - parametres }
  TobParametresAssocPopul := Tob.create('table d''association Population - parametres', nil, -1);
  TobParametresAssocPopul.LoadDetailDBFromSQL('PGPARAMETRESASSOC','SELECT PGO_CODEASSOCIE, PGO_PGPARAMETRE, PGO_DATEVALIDITE FROM PGPARAMETRESASSOC WHERE ##PGO_PREDEFINI## AND PGO_TYPEPARAMETRE = "POP" AND PGO_TYPEPOP = "'+TYP_POPUL_PRES+'"');
end;

procedure TGestionPresence.ReloadParametresPopul;
begin
  { Chargement de la table des parametres population }
  TobParametresPopul := Tob.create('table des parametres population', nil, -1);
  TobParametresPopul.LoadDetailDBFromSQL('PGPARAMETRES','SELECT PGP_PGPARAMETRE, PGP_LIBELLE, PGP_PGNOMCHAMP, PGP_PGVALCHAMP FROM PGPARAMETRES WHERE ##PGP_PREDEFINI## AND PGP_TYPEPARAMETRE = "POP" AND PGP_MODULECHAMP = "'+TYP_POPUL_PRES+'"');
end;

procedure TGestionPresence.ReloadProfilPresence;
begin
  TobProfilPresence := TOB.Create('Profils de présence', nil, -1);
  TobProfilPresence.LoadDetailFromSQL('SELECT PPQ_PROFILPRES, PPQ_DATEVALIDITE, PPQ_LIBELLE, '
                                     +' PPQ_TYPEPROFILPRES, PPQ_TPSTRAVAIL, '
                                     +' PPQ_LIMITEHTEMODUL, PPQ_LIMITEBASMODUL, PPQ_TYPSEUILT1, '
                                     +' PPQ_SEUILT1, PPQ_TYPSEUILT2, PPQ_SEUILT2, PPQ_TYPSEUILT3, '
                                     +' PPQ_SEUILT3, PPQ_NBJOURSCYCLE, PPQ_TYPCONTINGHSUP, '
                                     +' PPQ_CONTINGHSUP, PPQ_TYPEPLAFHREAN, PPQ_PLAFHREANNU, '
                                     +' PPQ_DATEREGANNU, '
                                     +' PPQ_DATEDEBUTAFF, PPQ_TYPEAFFECT, PPQ_CYCLEAFFECT '
                                     +', PPQ_JOURRTTLIBRE, PPQ_JOURRTTEMPL '
                                     +' FROM PROFILPRESENCE '
                                     +' ORDER BY PPQ_DATEVALIDITE');
end;

procedure TGestionPresence.ReloadProfilPresenceSalarie;
var
  stSQL : String;
begin
  TobProfilPresenceSalarie := TOB.Create('Profils de présences affectés au salarié', nil, -1);
  stSQL := 'SELECT PPZ_SALARIE, PPZ_DATEVALIDITE, '
         + ' PPZ_TYPPROFILPRES, PPZ_PROFILPRES '
         + ' FROM PROFILPRESSALARIE ';
  AddFiltreSalarie('PPZ', stSQL);
  stSQL := stSQL + ' ORDER BY PPZ_DATEVALIDITE';
  TobProfilPresenceSalarie.LoadDetailFromSQL(stSQL);
end;

procedure TGestionPresence.ReloadSalariePopul (Salarie : String='');
var
  stSQL : String;
begin
  { Chargement de la table d'association Salarié - Population }
  TobSalariePopul := Tob.create('table d''association Salarié - Population', nil, -1);
  stSQL := 'SELECT PNA_SALARIE, PNA_POPULATION FROM SALARIEPOPUL WHERE (PNA_TYPEPOP = "'+TYP_POPUL_PRES+'")';
  AddFiltreSalarie('PNA', stSQL);
  If Salarie <> '' Then StSQL := StSQL + ' AND PNA_SALARIE="'+Salarie+'"';
  TobSalariePopul.LoadDetailDBFromSQL('SALARIEPOPUL', stSQL);
end;

{ décomposition des rythmes en modèles de cycles }
function TGestionPresence.GetDecompositionModelesCyclesSalarie(TobCyclesSalarie: tob; DateDebutMin, DateFinMax : TDateTime): tob;    //Salarie: String;
var
  TobCycle : Tob;
  indexModele, indexCycle : Integer;
  temptob, temptobModele, tempTobDureeCycle, resultfille : Tob;
  DateValiditeCycle : Integer;
  CurDate, debCyc, FinCyc : TDateTime;
  Cycle : String;
  Modele, Journee : String;
  DateLancementCycle, DateDebutModele, DateFinModele : TDateTime;
  resteJours : Integer;
begin
  result := tob.Create('Modèles de cycle',nil,-1);
  for indexCycle := TobCyclesSalarie.FillesCount(0)-1 downto 0 do
  begin
    TobCycle := TobCyclesSalarie.Detail[indexCycle];
    if Assigned(TobCycle) and (TobCycle.GetString('TYPE') = 'CYC') then
    begin
      { On récupère les infos nécéssaires concernant le cycle }
      debCyc := TobCycle.GetDateTime('DATEDEBUT');
      if debCyc < DateDebutMin then debCyc := DateDebutMin;
      FinCyc := TobCycle.GetDateTime('DATEFIN');
      if FinCyc > DateFinMax then FinCyc := DateFinMax;
      Cycle :=  TobCycle.GetString('CYCLE');
      { la durée du cycle }
      tempTobDureeCycle := TobDureeCycle.FindFirst(['PYC_CYCLE'],[Cycle],False);
      if not Assigned(tempTobDureeCycle) or (tempTobDureeCycle.GetString('DUREE') = '') then exit;
      DateValiditeCycle := tempTobDureeCycle.GetInteger('DATE');
      CurDate := debCyc;
      if Trunc(CurDate) < DateValiditeCycle then CurDate := DateValiditeCycle;{ le cycle n'est pas valide avant sa date de validité }
      { On parcours les modèles du cycle jusqu'à la fin de la période }
      While CurDate <= FinCyc do
      begin { On recherche le modèle pour la date "CurDate" }
        GetModeleAndJourneeInCycle(Modele, Journee, DateLancementCycle, DateDebutModele,
                                   DateFinModele, resteJours, CurDate, Cycle, 'CYC', False);
        if Modele <> '' then
        begin
          temptobModele := GetAllValidesFromTobValidite(TobModeleCycle, ['PMY_MODELECYCLE'], [Modele], 'PMY_DATEVALIDITE',
                                                          'DATEDEBUT', DateDebutModele, 'DATEFIN', DateFinModele);
          for indexModele := temptobModele.FillesCount(0)-1 downto 0 do
          begin
            temptob := temptobModele.detail[indexModele];
            temptob.AddChampSupValeur('DATEVALIDITE', temptob.GetString('PMY_DATEVALIDITE'));
            temptob.AddChampSupValeur('CYCLE', Modele);
            temptob.AddChampSupValeur('TYPE', 'MOD');
            temptob.AddChampSupValeur('LIBELLE', temptob.GetString('PMY_LIBELLE'));
            temptob.ChangeParent(result, -1);
          end;
          FreeAndNil(temptobModele);
          CurDate := DateFinModele+1;
        end else
          CurDate := CurDate+1;
      end;
    end
    else // cas d'un modèle de cycle
    begin
     resultfille := TOB.Create('Modèle cycle', result, -1);
     resultfille.Dupliquer(TobCyclesSalarie.Detail[indexCycle], True, True);
    end;
  end;
end;

Procedure TGestionPresence.GetInfoPresenceSalarie(LaDate: TDateTime;
  Salarie: String; var Profil, Cycle, Modele, JourneeType: String;
  var isExceptionSalarie, isExceptionCycle: Boolean;
  CheckException: Boolean);
var
  TempTob : Tob;
  CycleOuModele, TypeDeCycle : String;
  DateLancementCycle, DateDebutModele, DateFinModele : TDateTime;
  resteJours : Integer;
begin
  isExceptionSalarie := False;
  isExceptionCycle := False;
  JourneeType := '';
  { Recherche des exceptions de présence du salarié }
  if CheckException then
  begin
    GetExceptionSalarie( LaDate, Salarie, CycleOuModele, TypeDeCycle);
    isExceptionSalarie := True;
  end;
  if TypeDeCycle = 'JOU' then
  begin
    JourneeType := CycleOuModele;
    exit;
  end;
  if (JourneeType = '') and (CycleOuModele = '') then
  begin
    { Recherche du profil de présence du salarié }
    isExceptionSalarie := False;
    Profil := GetProfil (Salarie, LaDate);
    
    { lecture du profil de présence du salarié valide à la date de recherche }
    TempTob := GetLastValideFromTob(LaDate,'PPQ_DATEVALIDITE', TobProfilPresence, ['PPQ_PROFILPRES'], [Profil]);
    if Assigned(TempTob) then
    begin
      CycleOuModele := TempTob.GetString('PPQ_CYCLEAFFECT');
      TypeDeCycle   := TempTob.GetString('PPQ_TYPEAFFECT');
    end;
  end;
  { Recherche de la journée type en fonction du cycle ou modèle }
  if (JourneeType = '') and (CycleOuModele <> '')then
  begin
    if TypeDeCycle = 'CYC' then
    begin
      Cycle := CycleOuModele;
      GetModeleAndJourneeInCycle(Modele, JourneeType, DateLancementCycle, DateDebutModele, DateFinModele, resteJours, LaDate, CycleOuModele, TypeDeCycle, CheckException)
    end else if TypeDeCycle = 'MOD' then
    begin
      Modele := CycleOuModele;
      JourneeType := GetJourneeType(LaDate, CycleOuModele, TypeDeCycle, CheckException);
    end;
    isExceptionCycle := DateOfException(LaDate, CycleOuModele, TypeDeCycle);
  end;
end;

function TGestionPresence.GetFiltreSalarie(const Prefixe : String) : String;
begin
  if FFiltreSalarie = '' then
    result := ''
  else begin
    if Prefixe = 'PSA' then
      result := StringReplace(FPSAFiltreSalarie, '###_', Prefixe+'_', [rfReplaceAll, rfIgnoreCase])
    else
      result := StringReplace(FFiltreSalarie, '###_', Prefixe+'_', [rfReplaceAll, rfIgnoreCase]);
  end;
end;

procedure TGestionPresence.SetFiltreSalarie(const Prefixe, Value: String);
begin
  if pos('PSA_', StringReplace(Value, 'PSA_SALARIE', '', [rfReplaceAll, rfIgnoreCase])) > 0 then
    FPSAFiltreSalarie := StringReplace(Value, Prefixe+'_', '###_', [rfReplaceAll, rfIgnoreCase])
  else begin
    FFiltreSalarie := StringReplace(Value, Prefixe+'_', '###_', [rfReplaceAll, rfIgnoreCase]);
    FPSAFiltreSalarie := FFiltreSalarie;
  end;
  { Suppression d'un éventuel WHERE dans le filtre : }
  FFiltreSalarie := UpperCase(Trim(StringReplace(FFiltreSalarie, 'WHERE', '', [rfReplaceAll, rfIgnoreCase])));
  FPSAFiltreSalarie := UpperCase(Trim(StringReplace(FPSAFiltreSalarie, 'WHERE', '', [rfReplaceAll, rfIgnoreCase])));
  { Ajout d'un AND au début, si nécéssaire }
  if LeftStr(FFiltreSalarie, 3)    <> 'AND' then
    FFiltreSalarie    := ' AND ('+FFiltreSalarie+')'
  else
    FFiltreSalarie    := ' '+FFiltreSalarie;
  if LeftStr(FPSAFiltreSalarie, 3) <> 'AND' then
    FPSAFiltreSalarie := ' AND ('+FPSAFiltreSalarie+')'
  else
    FPSAFiltreSalarie := ' '+FPSAFiltreSalarie;
end;


procedure TGestionPresence.AddFiltreSalarie(const Prefixe: String;
  var StQuery: String);
begin
  if Pos('WHERE', UpperCase(StQuery)) > 0 then
    StQuery := StQuery + GetFiltreSalarie(Prefixe)
  else
    StQuery := StQuery + ' WHERE ' + GetFiltreSalarie(Prefixe);
end;

(*//PT2
Procedure TGestionPresence.EclateTob(var FromTob, ToTob : Tob; ChampCritere : String; Duplicate : Boolean = True);  //; NoDoublons : Boolean = False
var
  indexTob : Integer;
  TempTob, FilleTob, PetiteFilleTob : Tob;
  TempValue : String;
  NomTob, NomTobPetiteFille : String;
//  function FindTobInTob(SearchTob, InTob : Tob) : Boolean;
//  var
//    index, indexInTob : Integer;
//    TempInTob : Tob;
//    NomChamp : String;
//  begin
//    result := False;
//    { On parcours les tobs filles de la tob dans laquelle on effectue la recherche }
//    for indexInTob := 0 to InTob.detail.count -1 do
//    begin
//      TempInTob := InTob.detail[indexInTob];
//      { Pour chaque champs (réel) de la tob à chercher }
//      result := True;
//      for index := 1 to SearchTob.NombreChampReel do
//      begin
//        { On compare la valeur du champs avec la valeur du même champs dans la fille courante }
//        NomChamp := SearchTob.GetNomChamp(index);
//        if not InTob.FieldExists(NomChamp) then
//        begin
//          result := False;
//          exit;
//        end;
//        if TempInTob.GetValue(NomChamp) <> SearchTob.GetValue(NomChamp) then
//        begin
//          result := False;
//          break;  //On passe à la fille suivante
//        end;
//        //Sinon on passe au champs suivant
//      end;
//      if result then exit;
//    end;
//  end;
begin
  if not Assigned(ToTob) then
    ToTob := Tob.create(FromTob.NomTable, nil, -1);
  NomTob := ToTob.NomTable;
  if FromTob.Detail.Count >= 1 then
    NomTobPetiteFille := FromTob.Detail[0].NomTable
  else
    NomTobPetiteFille := '';
  for indexTob := FromTob.Detail.Count -1 downto 0 do
  begin
    TempTob := FromTob.Detail[indexTob];
    TempValue := TempTob.GetString(ChampCritere);
    FilleTob := ToTob.FindFirst([ChampCritere], [TempValue], False);
    if not Assigned(FilleTob) then
    begin
      FilleTob := TOB.Create(NomTob, ToTob, -1);
      FilleTob.AddChampSupValeur(ChampCritere, TempValue);
    end;
//    if (not NoDoublons) or (NoDoublons and (not FindTobInTob(TempTob, FilleTob))) then
//    begin
    if Duplicate then
    begin
      PetiteFilleTob := TOB.Create(NomTob, FilleTob, -1);
      PetiteFilleTob.Dupliquer(TempTob, True, True, True);
    end else begin
      TempTob.ChangeParent(FilleTob, -1);
    end;
//    end;
  end;
end;
*)

function TGestionPresence.GetTobAbsencePresenceFiltreSalarie(Salarie: string): Tob;
begin
  if Assigned(TobAbsencePresenceSalarie) then
    result := TobAbsencePresenceSalarie.FindFirst(['PCN_SALARIE'], [Salarie], False)
  else                    
    result := Nil;
end;

{ TCalcuVarPre }

function TCalcuVarPre.EvalVarPreCompPres(Salarie, CompteurPresence: string; Perio : String; const DateDeb,
  DateFin: TDateTime; TOB_Rub : Tob; var NeedSave, HadRegul : Boolean; var Indicateur : String; Diag: TObject = nil; ForceRecalcul : Boolean = False): double;
var
  TempTob : Tob;
  Cptr : Integer;
begin
  NeedSave := True;
  HadRegul := False;
  Indicateur := 'CAL';
  result := 0;
  if not ForceRecalcul then
  begin
    Cptr := 0;
    TempTob := TobValeursCompteurs.FindFirst(['PYP_SALARIE'], [Salarie], False);
    if Assigned(TempTob) then
      TempTob := TempTob.FindFirst(['PYP_COMPTEURPRES'], [CompteurPresence], False);
    if Assigned(TempTob) then
    begin
      { On fait une boucle pour trouver les compteurs déjà calculés et leurs éventuelles régul }
      TempTob := TempTob.FindFirst(['PYP_PERIODICITEPRE', 'PYP_DATEDEBUTPRES', 'PYP_DATEFINPRES'], [Perio, DateDeb, DateFin], False);
      if Assigned(TempTob) then
      begin
        NeedSave := False;
        while Assigned(TempTob) do
        begin
          Inc(Cptr);
          result := result + TempTob.GetDouble('PYP_QUANTITEPRES');
          Indicateur := TempTob.GetString('PYP_PGINDICATPRES');
          if Cptr > 1 then HadRegul := True;
          TempTob := TempTob.FindNext(['PYP_PERIODICITEPRE', 'PYP_DATEDEBUTPRES', 'PYP_DATEFINPRES'], [Perio, DateDeb, DateFin], False);
        end;
      end;
    end;
  end;
  if NeedSave then
  begin
    { Si on ne trouve pas le compteurs, c'est qu'il n'a pas encore été calculé.
      Si c'est le cas, on récupère la valeur de la variable associé au compteur }
    TempTob := GetLastValideFromTob(DateFin, 'PYR_DATEVALIDITE', TobCompteurs, ['PYR_COMPTEURPRES'], [CompteurPresence]);
    if Assigned(TempTob) then
      result :=  ValVariable(TempTob.GetString('PYR_VARIABLEPRES'), DateDeb, DateFin, TOB_Rub, Diag);
  end;
end;


Procedure TCalcuVarPre.ConcatRegul(Salarie, CompteurPresence: string; Perio : String; const DateDeb,
  DateFin: TDateTime);
var
  TempTob, TobCompteur, TobRegul : Tob;
begin
  TempTob := TobValeursCompteurs.FindFirst(['PYP_SALARIE'], [Salarie], False);
  if Assigned(TempTob) then
    TempTob := TempTob.FindFirst(['PYP_COMPTEURPRES'], [CompteurPresence], False);
  if Assigned(TempTob) then
  begin
    { On fait une boucle pour trouver les compteurs déjà calculés et leur éventuel régul }
    TempTob := TempTob.FindFirst(['PYP_PERIODICITEPRE', 'PYP_DATEDEBUTPRES', 'PYP_DATEFINPRES'], [Perio, DateDeb, DateFin], False);
    TobCompteur := TempTob;
    while Assigned(TempTob) do
    begin
      if TempTob.GetDateTime('PYP_DATEDEBUTBUL') < TobCompteur.GetDateTime('PYP_DATEDEBUTBUL') then
      begin
        TobRegul := TobCompteur;
        TobCompteur := TempTob;
      end else begin
        TobRegul := TempTob;
      end;
      if Assigned(TobRegul) then
      begin
        TobCompteur.PutValue('PYP_QUANTITEPRES',TobCompteur.GetDouble('PYP_QUANTITEPRES') + TobRegul.GetDouble('PYP_QUANTITEPRES'));
        TobCompteur.PutValue('PYP_DATEFINBUL',TobRegul.GetDateTime('PYP_DATEFINBUL'));
        TobRegul.DeleteDB();
        TobRegul.ChangeParent(nil, -1);
        FreeAndNil(TobRegul);
      end;
      TempTob := TempTob.FindNext(['PYP_PERIODICITEPRE', 'PYP_DATEDEBUTPRES', 'PYP_DATEFINPRES'], [ Perio, DateDeb, DateFin], False);
    end;
  end;
end;

function TCalcuVarPre.EvalVarPreJType(Salarie, Droit: string; const DateDeb,
  DateFin: TDateTime; TypeBase : Integer; Diag: TObject = nil; JourneeType : String = 'A Rechercher'): double;
var
  TempTob : Tob;
  curDate : TDateTime;
  CycleouModele, typeCycle : String;
  isExceptionSal, isExceptionCycle : Boolean;
begin
  result := 0;
  // Boucle sur les journées types du salarié pour la période.
  curDate := DateDeb;
  While curDate <= DateFin do
  begin
    if JourneeType = 'A Rechercher' then
      JourneeType := GetJourneeTypeSalarie(curDate, Salarie, CycleouModele, typeCycle, isExceptionSal, isExceptionCycle);
    // Somme les droits correspondants
    TempTob := TobDroitsJourneesTypes.FindFirst(['PDJ_JOURNEETYPE', 'PDJ_PGDROIT'], [JourneeType, Droit], False);
    if Assigned(TempTob) then
    begin
      if TypeBase = 506 then
        result := result + TempTob.GetDouble('PDJ_QUANTITE1');
      if TypeBase = 507 then
        result := result + TempTob.GetDouble('PDJ_QUANTITE2');
      if TypeBase = 508 then
        result := result + TempTob.GetDouble('PDJ_QUANTITE1') + TempTob.GetDouble('PDJ_QUANTITE2');
    end;
    curDate := IncDay(curDate);
  end;
end;

function TCalcuVarPre.EvalVarPreEvt(Salarie,MotifEvenement : string; const ResultColonneName: String;
  const DateDeb, DateFin: TDateTime; Diag: TObject = nil): double;
var
  TempTob, TobAbsPresSalarie : Tob;
  dateDebutEvt, DateFinEvt : TDateTime;
  dateheure : String;
  indexAbsPre : Integer;
begin
  result := 0;
  //Boucle sur les evenements du salarié pour la période.
  //On récupère les absences et présences du salarié
  TobAbsPresSalarie := GetTobAbsencePresenceFiltreSalarie(Salarie);
  //On filtre le motif recherché
  if assigned(TobAbsPresSalarie) then
  begin
    for indexAbsPre := 0 to TobAbsPresSalarie.Detail.Count -1 do
    begin
      TempTob := TobAbsPresSalarie.Detail[indexAbsPre];
      if Assigned(TempTob) and (TempTob.GetString('PCN_TYPECONGE') = MotifEvenement) then
      begin
        //On filtre la période de dates recherchée
        dateDebutEvt := TempTob.GetDateTime('PCN_DATEDEBUTABS');
        DateFinEvt := TempTob.GetDateTime('PCN_DATEFINABS');
        if (DateDeb <= dateFinEvt) and (DateFin >= dateDebutEvt) then
        begin
          if ResultColonneName = 'QTE' then      // quantité
          begin
            if Temptob.Getstring('PMA_JOURHEURE') = 'QTE' then
              result := result + Temptob.GetDouble('PCN_HEURES');
          end else if (ResultColonneName = 'PLAGE1')  then   // Evt sur la plage 1
          begin
            result := result + Min(DateFinEvt, DateFin) - Max(dateDebutEvt, DateDeb);
            if not (    ((dateDebutEvt = DateFinEvt) or (dateDebutEvt = DateFin))
                    and (TempTob.GetString('PCN_DEBUTDJ') = 'PAM')               ) then
              result := result + 1;
          end else if (ResultColonneName = 'PLAGE2') then    // Evt sur la plage 2
          begin
            result := result + Min(DateFinEvt, DateFin) - Max(dateDebutEvt, DateDeb);
            if not (    ((dateDebutEvt = DateFinEvt) or (dateFinEvt = DateDeb))
                //    and (TempTob.GetString('PCN_DEBUTDJ') = 'MAT')               ) then       pt4
                      and (TempTob.GetString('PCN_FINDJ') = 'MAT')               ) then       //pt4
              result := result + 1;
          end else if (ResultColonneName = 'JOURS') then    // Evt sur la journée complète
          begin
            if (TempTob.Getstring('PMA_JOURHEURE') = 'JOU') then      // si evt type Jour
            begin
              { On compte le nombre de jour de l'évenement qui font parti de la période de calcul }
              result := result + Min(DateFinEvt, DateFin) - Max(dateDebutEvt, DateDeb) +1;
              { Si l'évènement dure sur plusieurs jours }
              if (dateDebutEvt <> DateFinEvt) then
              begin
                { Si le premier jour de l'évènement fait parti de la période de calcul
                  on enlève la journée si elle début seulement l'après-midi }
                if (    ((dateDebutEvt >= DateDeb) and (dateDebutEvt <= DateFin))
                    and (TempTob.GetString('PCN_DEBUTDJ') = 'PAM') ) then
                  result := result - 1;
                { Si le dernier jour de l'évènement fait parti de la période de calcul
                  on enlève la journée si elle fini le matin }
                if (    ((DateFinEvt >= DateDeb) and (DateFinEvt <= DateFin))
                    and (TempTob.GetString('PCN_FINDJ') = 'MAT') ) then
                  result := result - 1;
              { Si l'évènement dure sur un seul jour }
              end else begin 
                { Si l'évènement ne concerne qu'une demi-journée, on ne le compte pas }
                if not ((TempTob.GetString('PCN_DEBUTDJ') = 'MAT') and (TempTob.GetString('PCN_FINDJ') = 'PAM')) then
                  result := result - 1;
              end;

              { TODO 3 -oGGU -cPrésence : Gestion des évènements avec heure de début et heure de fin (en horaire) }

            end;
          end else begin  // PCN_HEURES ou PCN_NBHEURESNUIT
           if (ResultColonneName = 'PCN_HEURES') or (ResultColonneName = 'PCN_NBHEURESNUIT') then
             result := result + TempTob.GetDouble(ResultColonneName)
           else
           begin // PCN_HDEB ou PCN_HFIN
             Dateheure := TempTob.Getstring(ResultColonneName);
             DateHeure := Copy(DateHeure, 1 ,5);
             result := strTimeToFloat(Dateheure,True);
           end;
          end;
        end;
      end;
    end;
  end;
end;

{ Divise les compteurs à calculer en fonction de leur périodicité }
Procedure TCalcuVarPre.Traiteperiodicite(var TobCompteursACalculer: tob;
  Salarie : String; DateEntreeSalarie, DateSortieSalarie : TDateTime;
//  CalculJour, CalculHebdo, CalculMens, CalculAnn, CalculFinCycle : Boolean;
  DateDebutJour, DateDebutHebdo, DateDebutMens, DateDebutAnn, DateDebutFinCycle,
  DateFinJour, DateFinHebdo, DateFinMens, DateFinAnn, DateFinFinCycle : TDateTime);
var
  indexCompteur, indexListeMois : Integer;
  TobCompteur, StockNewTob, tempTobDureeCycle : Tob;
  DureeCycle : Integer;
  periodicite : String;
//  periodeRaz  : String;
  Modele, Journee, CycleOuModele, TypeDeCycle : String;
  DateLancement, DateDebutModele, DateFinModele : TDateTime;
  ResteJours : Integer;
  isExceptionSalarie, isExceptionCycle : Boolean;
  debut, fin, tpDeb, tpFin, tpNewDeb, tpNewFin  : TDateTime; 
  procedure SupprTob(LaTob : Tob);
  begin
    LaTob.ChangeParent(nil,-1);
    FreeAndNil(LaTob);
  end;
  Procedure DiviseTob(LaTob, ToTob : Tob; NewDateDebut, NewDateFin: TDateTime);
  var
    NewTob : Tob;
  begin
    NewTob := tob.Create(LaTob.NomTable,ToTob, -1);
    NewTob.Dupliquer(LaTob, True, True);
    NewTob.PutValeur(idDATEDEBUT, NewDateDebut);
    NewTob.PutValeur(idDATEFIN, NewDateFin);
  end;
  Function ValidePeriode(const DebutCalcul, FinCalcul, DebutPeriode, FinPeriode, SortieSalarie : TDateTime) : Boolean;
  var
    FinPeriodeDansCalcul, FinPeriodeDansPresence, SortieDansPeriode : Boolean;
  begin
    { Assertion : Le salarié est forcément rentré avant le début de la période de calcul
     ( if DateEntreeSalarie > 10 then debut := max(TobCompteur.GetDateTime(idDATEDEBUT), DateEntreeSalarie)
       else debut := TobCompteur.GetDateTime(idDATEDEBUT); )    }
    { La période à tester se termine dans la période de calcul }
    FinPeriodeDansCalcul := (DebutCalcul <= FinPeriode) and (FinPeriode <= FinCalcul);
    { Le salarié est présent quand la période à tester se termine }
    FinPeriodeDansPresence := (DebutCalcul <= FinPeriode) and ((SortieSalarie < 10) or (FinPeriode <= SortieSalarie));
    if (not FinPeriodeDansCalcul) then
      result := False
    else if FinPeriodeDansPresence then
      result := True
    else begin
      { Le salarié à quitté l'entreprise dans la période à tester }
      SortieDansPeriode := (DebutPeriode <= SortieSalarie) and (SortieSalarie <= FinPeriode);
      result := SortieDansPeriode;
    end;
  end;
begin
  StockNewTob := TOB.Create('',nil,-1);
  if (not IsNumChampSaved) and (TobCompteursACalculer.Detail.Count >= 1)  then
  begin
    idDATEDEBUT          := TobCompteursACalculer.Detail[0].GetNumChamp('DATEDEBUT');
    idDATEFIN            := TobCompteursACalculer.Detail[0].GetNumChamp('DATEFIN');
    idPYR_PERIODICITEPRE := TobCompteursACalculer.Detail[0].GetNumChamp('PYR_PERIODICITEPRE');
    IsNumChampSaved := True;
//    idPYR_PERIODERAZ   := TobCompteursACalculer.Detail[0].GetNumChamp('PYR_PERIODERAZ');
  end;
  for indexCompteur := TobCompteursACalculer.FillesCount(0)-1 downto 0 do
  begin

    TobCompteur := TobCompteursACalculer.Detail[indexCompteur];
    periodicite := TobCompteur.GetString(idPYR_PERIODICITEPRE);
//    periodeRaz  := TobCompteur.GetString(idPYR_PERIODERAZ);
    if DateEntreeSalarie > 10 then
      debut := max(TobCompteur.GetDateTime(idDATEDEBUT), DateEntreeSalarie)
    else
      debut := TobCompteur.GetDateTime(idDATEDEBUT);
    fin := TobCompteur.GetDateTime(idDATEFIN);

    if CalculAnn and (periodicite = periodiciteAnnuelle) then { Annuelle }
    begin
      { On ne prends que les compteurs valides en fin d'année }
      if ValidePeriode(debut, fin, DateDebutAnn, DateFinAnn, DateSortieSalarie) then
      begin
        DiviseTob(TobCompteur, StockNewTob, DateDebutAnn, DateFinAnn);
      end;
    end else if (periodicite = periodiciteMensuelle) then { Mensuelle }
    begin
      { Récupérer les dates de fin de mois correspondantes }
      indexListeMois := 0;
      While length(ListeFinDeMois) >= indexListeMois + 1 do
      begin
        if IncDay(ListeFinDeMois[indexListeMois]) <= DateDebutMens then
          break;
        Inc(indexListeMois);
      end;
      While (indexListeMois >= 1) do
      begin
        tpDeb := IncDay(ListeFinDeMois[indexListeMois]);
        tpFin := ListeFinDeMois[indexListeMois-1];
        { On ne prends que les compteurs valides en fin de mois
          et le compteur du mois de départ du salarié (iLast = 1) }
        if ValidePeriode(debut, fin, tpDeb, tpFin, DateSortieSalarie) then
        begin
          if (tpDeb < tpFin) then DiviseTob(TobCompteur, StockNewTob, tpDeb, tpFin);
        end;
        Dec(indexListeMois);
      end;
    end else if (periodicite = periodiciteHebdomadaire) then
    begin { Hebdomadaire -> Calcul pour toutes les semaines terminées dans la période }
      { On début au début de la semaine }
      tpDeb := debut - JourDeLaSemaine(debut) + 1; 
      { On fini en fin de semaine complète }
      if jourDeLaSemaine(fin) <> 7 then
      tpFin := fin - JourDeLaSemaine(fin)
      else tpfin := fin ;
      While tpDeb < tpFin do
      begin
        tpNewDeb := tpDeb;
        tpNewFin := tpDeb + 6;
        if ValidePeriode(debut, fin, tpNewDeb, tpNewFin, DateSortieSalarie) then
          DiviseTob(TobCompteur, StockNewTob, tpNewDeb, tpNewFin);
        tpDeb := tpNewFin + 1;
      end;
    end else if (periodicite = periodiciteJournaliere) then
    begin { Journalière -> Calcul pour chaque jour de la période }
      tpDeb := debut;
      if DateSortieSalarie > 10 then
        tpFin := min(fin, DateSortieSalarie)
      else
        tpFin := fin;
      While tpDeb <= tpFin do
      begin
        tpNewDeb := tpDeb;
        tpNewFin := tpDeb;
        if ValidePeriode(debut, fin, tpNewDeb, tpNewFin, DateSortieSalarie) then
          DiviseTob(TobCompteur, StockNewTob, tpNewDeb, tpNewFin);
        tpDeb := tpNewFin + 1;
      end;
    end else if (periodicite = periodiciteFinDeCycle) then
    begin { Fin de cycle }
      tpDeb := debut;
      tpFin := fin;
      While tpDeb <= tpFin do
      begin
        GetJourneeTypeSalarie(tpDeb, Salarie, CycleOuModele, TypeDeCycle, isExceptionSalarie, isExceptionCycle);
        if TypeDeCycle = 'CYC' then
        begin
          tempTobDureeCycle := TobDureeCycle.FindFirst(['PYC_CYCLE'],[CycleOuModele],False);
          if (not Assigned(tempTobDureeCycle)) or (tempTobDureeCycle.GetString('DUREE') = '') then
          begin
            tpDeb := tpDeb +1;
            continue;
          end;
          DateLancement := tempTobDureeCycle.GetDateTime('DATE');
          DureeCycle := tempTobDureeCycle.GetInteger('DUREE');
          { le cycle n'est pas valide avant sa date de validité }
          if Trunc(tpDeb) < Trunc(DateLancement) then
          begin
            tpDeb := tpDeb +1;
            continue;
          end;
          DateDebutModele := DateLancement + DureeCycle * Trunc((Trunc(tpDeb) - Trunc(DateLancement))/DureeCycle);
          DateFinModele := DateDebutModele + DureeCycle - 1;
        end else begin
          GetModeleAndJourneeInCycle(Modele, Journee, DateLancement, DateDebutModele, DateFinModele, ResteJours, tpDeb, CycleOuModele, TypeDeCycle);
        end;
        tpNewDeb := DateDebutModele;
        tpNewFin := DateFinModele;
        { On ne prends que les cycles terminés dans la période }
        if ValidePeriode(debut, fin, tpNewDeb, tpNewFin, DateSortieSalarie) then
          DiviseTob(TobCompteur, StockNewTob, tpNewDeb, tpNewFin);
        tpDeb := tpNewFin + 1;
      end;
    end;
  end;
  TobCompteursACalculer.ClearDetail;
  TobCompteursACalculer.Dupliquer(StockNewTob, True, True);
  FreeAndNil(StockNewTob);
end;


constructor TCalcuVarPre.Create(FiltreSalarie: String = ''; DateDebut : TDateTime = 0; DateFin : TDateTime = 0; DateFinEstFinDeMois : Boolean = False);
var
  stSQL : String;
  TempTob, TempTobNewFille, TempTobVarPai : Tob;
  ListeParam, Param, perio : String;
  q: TQuery;
  indexSalarie, indexVarPai, IndexDate : Integer;
  FindInListe : Boolean;
  TempInverseListeFinDeMois : Array of TDateTime;
  indexInversion : Integer;
begin
  if DateDebut = 0 then DateDebut := Date()-366;
  if DateFin   = 0 then DateFin   := iDate2099;
  IsNumChampSaved := False;
  if FiltreSalarie <> '' then SetFiltreSalarie('PSA', FiltreSalarie);
  inherited Create(True,True,True,True,True,DateDebut,DateFin);
  TempTob := Tob.Create('Valeurs des compteurs de présence', nil, -1);
  { Recherche des compteurs utilisés dans les variables qui calcul des périodes précédantes à la période en cours }
  { Recherche des variables de paie faisant appel à des compteurs }
  TempTobVarPai := Tob.Create('variables de paie faisant appel à des compteurs', nil, -1);
  TempTobVarPai.LoadDetailFromSQL('select PVA_PERIODECALCUL, PVA_BASE0, pva_nbremoisgliss, pva_varperiodicite from variablepaie where ##pva_predefini## and pva_typevariable = "PRE" and PVA_TYPEBASE0 in ("500", "502", "504", "520")');
  { Parcours des variables de paie }
  { On fait la liste des variables de chaque PVA_PERIODECALCUL }
  for indexVarPai := 0 to TempTobVarPai.Detail.Count -1 do
  begin
    perio := TempTobVarPai.Detail[indexVarPai].GetString('PVA_PERIODECALCUL');
    if      (perio = '004') or (perio = '012') then
      ListeCompteurs412 := ListeCompteurs412 + ' "'+TempTobVarPai.Detail[indexVarPai].GetString('PVA_BASE0')+'",'
    else if (perio = '008') or (perio = '013') then
      ListeCompteurs813 := ListeCompteurs813 + ' "'+TempTobVarPai.Detail[indexVarPai].GetString('PVA_BASE0')+'",'
    else if (perio = '009') then
      ListeCompteurs9 := ListeCompteurs9     + ' "'+TempTobVarPai.Detail[indexVarPai].GetString('PVA_BASE0')+'",'
    else if (perio = '010') or (perio = '014') then
      ListeCompteurs1014 := ListeCompteurs1014   + ' "'+TempTobVarPai.Detail[indexVarPai].GetString('PVA_BASE0')+'",'
    else if (perio = '011') then
      ListeCompteurs11 := ListeCompteurs11   + ' "'+TempTobVarPai.Detail[indexVarPai].GetString('PVA_BASE0')+'",';
  end;
  FreeAndNil(TempTobVarPai);
  if ListeCompteurs412 <> '' then
    ListeCompteurs412 := LeftStr(ListeCompteurs412, Length(ListeCompteurs412) -1);
  if ListeCompteurs813 <> '' then
    ListeCompteurs813 := LeftStr(ListeCompteurs813, Length(ListeCompteurs813) -1);
  if ListeCompteurs9 <> '' then
    ListeCompteurs9 := LeftStr(ListeCompteurs9, Length(ListeCompteurs9) -1);
  if ListeCompteurs1014 <> '' then
    ListeCompteurs1014 := LeftStr(ListeCompteurs1014, Length(ListeCompteurs1014) -1);
  if ListeCompteurs11 <> '' then
    ListeCompteurs11 := LeftStr(ListeCompteurs11, Length(ListeCompteurs11) -1);

  stSQL := BuildLoadingQuery(DateDebut, DateFin);

//  stSQL := 'SELECT PYP_SALARIE, PYP_DATEDEBUTPRES, PYP_DATEFINPRES, PYP_COMPTEURPRES, '
//         + ' PYP_PERIODICITEPRE, PYP_THEMEPRE, PYP_QUANTITEPRES, PYP_ETATPRES, PYP_PGINDICATPRES, '
//         + ' PYP_TYPECALPRES, PYP_DATEDEBUTBUL, PYP_DATEFINBUL, PYP_DATECREATION, PYP_DATEMODIF FROM PRESENCESALARIE '
//         ;
//  stSQL := stSQL + 'WHERE (';
//  { On doit chargé tous les compteurs utilisés par les variables de paie "Date à Date" et "Depuis (hors période en cours)"
//    qui ont été calculés depuis le début de l'année }
//  if ListeCompteurs412 <> '' then
//    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs412+') and '+ConcernePeriode(EncodeDateDay(YearOf(DateDebut),1), DateFin)+')';
//  { On doit chargé tous les compteurs utilisés par les variables de paie "Mois en cours" et "Semaines précédentes du mois"
//    qui ont été calculés depuis le début du mois }
//  if ListeCompteurs813 <> '' then
//  begin
//    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
//    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs813+') and '+ConcernePeriode(EncodeDateMonthWeek(YearOf(DateDebut), MonthOf(DateDebut), 1 ,1), DateFin)+')';
//  end;
//  { On doit chargé tous les compteurs utilisés par les variables de paie "Semaine en cours"
//    qui ont été calculés depuis le début - 1 semaine }
//  if ListeCompteurs9 <> '' then
//  begin
//    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
//    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs9+') and '+ConcernePeriode(EncodeDateWeek(YearOf(DateDebut-7), WeekOf(DateDebut-7)), DateFin)+')';
//  end;
//  { On doit chargé tous les compteurs utilisés par les variables de paie "Mois glissants (hors mois en cours)"
//    qui ont été calculés depuis le début -12 mois }
//  if ListeCompteurs1014 <> '' then
//  begin
//    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
//    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs1014+') and '+ConcernePeriode(EncodeDateMonthWeek(YearOf(DateDebut)-1, MonthOf(DateDebut), 1 ,1), DateFin)+')';
//  end;
//  { On doit chargé tous les compteurs utilisés par les variables de paie "Semaine glissante (hors semaine en "
//    qui ont été calculés depuis le début -3 mois }
//  if ListeCompteurs11 <> '' then
//  begin
//    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
//    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs11+') and '+ConcernePeriode(EncodeDateMonthWeek(YearOf(IncMonth(DateDebut,-3)), MonthOf(IncMonth(DateDebut,-3)), 1 ,1), DateFin)+')';
//  end;
////  { TODO 3 -cPrésence -oGGU : Ajouter le chargement des compteurs mensuels susceptibles d'être passés en paie }
////  if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
////  stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in (select pyr_compteurpres from compteurpresence where pyr_integrepaie = "X") and '+ConcernePeriode(DateDebut, DateFin)+')';
//  if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
//  stSQL2 := stSQL2 + ConcernePeriode(DateDebut, DateFin);
//  stSQL := stSQL + stSQL2 + ')';
//  AddFiltreSalarie('PYP', stSQL);
  TempTob.LoadDetailDBFromSQL('PRESENCESALARIE', stSQL);
  { On éclate la tob des compteurs par salarié }
  EclateTob(TempTob, TobValeursCompteurs, 'PYP_SALARIE', False);
  FreeAndNil(TempTob);
  { On éclate la tob des compteurs par compteur }
  for indexSalarie := TobValeursCompteurs.Detail.Count -1 downto 0 do
  begin
    TempTob := TobValeursCompteurs.Detail[indexSalarie];
    TempTobNewFille := TOB.Create(TempTob.NomTable, nil, -1);
    EclateTob(TempTob, TempTobNewFille, 'PYP_COMPTEURPRES', False);
    TempTobNewFille.AddChampSupValeur('PYP_SALARIE', TempTob.GetString('PYP_SALARIE'));
    FreeAndNil(TempTob);
    TempTobNewFille.ChangeParent(TobValeursCompteurs, -1);
  end;
  TobDroitsJourneesTypes := Tob.Create('Droits associés à la journée type', nil, -1);
  TobDroitsJourneesTypes.LoadDetailFromSQL('SELECT PDJ_JOURNEETYPE, PDJ_PGDROIT, PDJ_QUANTITE1, PDJ_QUANTITE2, PDJ_LIBELLE, PDJ_DATEMODIF FROM DROITJOURNEETYPE');
  TobMotifsEvenements := Tob.Create('Tob des droits associés à la journée type', nil, -1);
  TobMotifsEvenements.LoadDetailFromSQL( 'SELECT PMA_MOTIFABSENCE, PMA_TYPERTT, PMA_RUBRIQUE, PMA_RUBRIQUEJ, '
                                       + ' PMA_ALIMENT, PMA_ALIMENTJ, PMA_GERECOMM, PMA_JOURHEURE, PMA_GESTIONMAXI, PMA_TYPEPERMAXI, '
                                       + ' PMA_PERMAXI, PMA_JRSMAXI, PMA_CALENDSAL, PMA_CALENDCIVIL, PMA_OUVRES, PMA_OUVRABLE, '
                                       + ' PMA_SSJOURFERIE, PMA_CONTROLMOTIF, PMA_PROFILABS, PMA_PROFILABSJ, PMA_TYPEATTEST, PMA_PREDEFINI, '
                                       + ' PMA_NODOSSIER, PMA_EDITION, PMA_PRISTOTAL, PMA_MOTIFEAGL, PMA_OKSAISIERESP, PMA_OKSAISIESAL, '
                                       + ' PMA_GESTIONIJSS, PMA_TYPEABS, '
                                       + ' PMA_EDITPLANPAIE, PMA_EDITPLANABS, PMA_ALIMNETH, '
                                       + ' PMA_ALIMNETJ, PMA_TYPEMOTIF, PMA_CTRLABSEXISTE, PMA_CTRLPREEXISTE, PMA_CTRLPLAGEH FROM MOTIFABSENCE');
  TobJourneesTypes := Tob.Create('Tob des journées types', nil, -1);
  TobJourneesTypes.LoadDetailFromSQL('SELECT PJO_JOURNEETYPE, PJO_LIBELLE, PJO_HORDEBPLAGE1, '
                                   + ' PJO_HORFINPLAGE1, PJO_DUREEPLAGE1, PJO_JOURJ1PLAGE1, PJO_TYPEJOUR1, '
                                   + ' PJO_HORDEBPLAGE2, PJO_HORFINPLAGE2, PJO_DUREEPLAGE2, PJO_JOURJ1PLAGE2, '
                                   + ' PJO_TYPEJOUR2, PJO_TYPEJOUR, PJO_EQUIVTPSPLEIN, PJO_JOURTRAVFERIE, '
                                   + ' PJO_POIDSJOUR, PJO_TEMPSLIBRE1, PJO_TEMPSLIBRE2, PJO_TEMPSLIBRE3, '
                                   + ' PJO_TEMPSLIBRE4, PJO_TEMPSLIBRE5, PJO_TEMPSLIBRE6, PJO_TEMPSLIBRE7, '
                                   + ' PJO_DUREEPAUSE, PJO_PAIEMENTPAUSE, PJO_PAUSEEFFECTIF, PJO_DUREENOTRAV, '
                                   + ' PJO_DUREENOTRAVEFF, PJO_ABREGE, PJO_COLORJOURTYPE'
                                   + ' FROM JOURNEETYPE ');
  TobCompteurs := Tob.Create('Compteurs ', nil, -1);
  TobCompteurs.LoadDetailFromSQL('SELECT PYR_PREDEFINI, PYR_NODOSSIER, PYR_COMPTEURPRES, PYR_DATEVALIDITE, '
                            + ' PYR_LIBELLE, PYR_PERIODICITEPRE, PYR_PERIODERAZ, PYR_INTEGREPAIE, PYR_RUBRIQUE, '
                            + ' PYR_TYPECHAMPRUB, PYR_VARIABLEPRES, PYR_EDITPLANPRES, PYR_PGCOLORPRE, PYR_THEMEPRE '
                            + ' FROM COMPTEURPRESENCE WHERE ##PYR_PREDEFINI## ');
  Tob_Ferie := ChargeTobFerie(DateDebut,DateFin);

  TobSaveCalcVar := Tob.Create('Sauvegarde des variables calculées', nil, -1);

  { Chargement du tableau des dates de fin de mois }
  SetLength(ListeFinDeMois, 0);
  { On ajoute la date de fin si on a coché le calcul des compteurs mensuels (DateFinEstFinDeMois à true)
    et que la liste de fin de mois ne contient pas déjà cette date
  }
  if DateFinEstFinDeMois then
  begin
    { Recherche de la date dans la liste }
    FindInListe := False;
    for IndexDate := 0 to Length(ListeFinDeMois) - 1 do
    begin
      if ListeFinDeMois[IndexDate] = Datefin then
      begin
        FindInListe := True;
        Break;
      end;
    end;
    if not FindInListe then
    begin
      SetLength(ListeFinDeMois, Length(ListeFinDeMois) + 1 );
      ListeFinDeMois[0] := Datefin;
    end;
  end;
  { On ajoute les dates de fin des mois cloturés (enregistrés dans les parametres société)
    Il faut inverser cette liste, parce qu'on veut la liste du plus recent au plus ancien
    alors qu'on stock la liste du plus ancien au plus recent
  }
  SetLength(TempInverseListeFinDeMois, 0);
  ListeParam := GetParamSocDatesCloture(MOIS);  //PT3 GetParamSocSecur(PresenceRenvoieParamSoc(MOIS), '');
  Param := ReadTokenSt(ListeParam);
  While (Param <> '') Do
  Begin
    SetLength(TempInverseListeFinDeMois, Length(TempInverseListeFinDeMois) + 1 );
    TempInverseListeFinDeMois[Length(TempInverseListeFinDeMois) - 1] := StrToDate(Param);
    Param := ReadTokenSt(ListeParam);
  End;
//  ListeParam := GetParamSocSecur(PresenceRenvoieParamSoc(MOIS), '');
//  Param := ReadTokenSt(ListeParam);
//  While (Param <> '') Do
//  Begin
//    SetLength(ListeFinDeMois, Length(ListeFinDeMois) + 1 );
//    ListeFinDeMois[Length(ListeFinDeMois) - 1] := StrToDate(Param);
//    Param := ReadTokenSt(ListeParam);
//  End;
  for indexInversion := (Length(TempInverseListeFinDeMois) -1) downto 0 do
  begin
    SetLength(ListeFinDeMois, Length(ListeFinDeMois) + 1 );
    ListeFinDeMois[Length(ListeFinDeMois) - 1] := TempInverseListeFinDeMois[indexInversion];
  end;

  { On recherche la date de début (la première date avec des compteurs journaliers) }
  SetLength(ListeFinDeMois, Length(ListeFinDeMois) + 1);
  q := OpenSQL('select min(pyp_datedebutpres) from presencesalarie where PYP_PGINDICATPRES <> "ARE" ', TRUE);
  if (not q.eof) and (q.Fields[0].AsDateTime > 10) then
    ListeFinDeMois[Length(ListeFinDeMois) -1] := q.Fields[0].AsDateTime - 1
  else  { si on ne trouve pas de date correspondant (aucun compteur calculé), on prends la date de début qui a été saisie }
    ListeFinDeMois[Length(ListeFinDeMois) -1] := DateDebut - 1;
  Ferme(q);
end;

destructor TCalcuVarPre.Destroy;
begin
  FreeAndNil(TobValeursCompteurs);
  FreeAndNil(TobDroitsJourneesTypes);
  FreeAndNil(TobMotifsEvenements);
  FreeAndNil(TobJourneesTypes);
  FreeAndNil(TobCompteurs);
  FreeAndNil(TobSaveCalcVar);
  SetLength(ListeFinDeMois, 0);
  if Assigned(Tob_Ferie) then FreeAndNil(Tob_Ferie);
  inherited;
end;

procedure TCalcuVarPre.GetCurrentMonth(var DateDebut, DateFin: TDateTime);
begin
  if Length(ListeFinDeMois) >= 2 then
  begin
    DateDebut := IncDay(ListeFinDeMois[1]);
    DateFin   := ListeFinDeMois[0];
  end else begin
    DateDebut := iDate2099;
    DateFin   := iDate2099;
  end;
end;

procedure TCalcuVarPre.GetCurrentWeek(var DateDebut, DateFin: TDateTime);
begin
  PresenceDonneSemaineCalculActuelle(DateDebut, DateFin);
//  { On début au début de la semaine }
//  DateDebut := Date() - JourDeLaSemaine(Date()) + 1;
//  { On fini en fin de semaine complète }
//  DateFin := DateDebut + 6;
end;

procedure TCalcuVarPre.GetMonthAbs(const decalage: Integer; var DateDebut,
  DateFin: TDateTime);
begin
  if Length(ListeFinDeMois) >= decalage + 2 then
  begin
    DateDebut := IncDay(ListeFinDeMois[decalage + 1]);
    DateFin := ListeFinDeMois[decalage];
  end else begin
    DateDebut := iDate2099;
    DateFin   := iDate1900;
  end;
end;

procedure TCalcuVarPre.GetMonth(const decalage: Integer; const reference : TDateTime; var DateDebut,
  DateFin: TDateTime);
var
  curDec : Integer;
  CurReference : TDateTime;
begin
  DateDebut := iDate2099;
  DateFin   := iDate1900;
  curDec := decalage;
  CurReference := reference;
  while curDec >= 0 do
  begin
    GetMonth(CurReference, DateDebut, DateFin);
    CurReference := DateDebut-1;
    Dec(curDec);
  end;
end;

procedure TCalcuVarPre.GetMonth(const Value: TDateTime; var DateDebut,
  DateFin: TDateTime);
var
  decalage : Integer;
  tpdeb, tpfin : TDateTime;
begin
  DateDebut := iDate2099;
  DateFin   := iDate1900;
  decalage := 0;
  While Length(ListeFinDeMois) >= decalage + 2 do
  begin
    GetMonthAbs(decalage, tpdeb, tpfin);
    if (tpdeb <= Value) and (Value <= tpfin) then
    begin
      DateDebut := tpdeb;
      DateFin   := tpfin;
      exit;
    end;
    Inc(decalage);
  end;
end;

procedure TCalcuVarPre.GetWeek(const Value: TDateTime; var DateDebut,
  DateFin: TDateTime);
begin
  { On début au début de la semaine }
  DateDebut := Value - JourDeLaSemaine(Value) + 1;
  { On fini en fin de semaine complète }
  DateFin := DateDebut + 6;
end;

procedure TCalcuVarPre.GetWeek(const decalage: Integer; var DateDebut,
  DateFin: TDateTime; const reference : TDateTime = 0);
var
  CurReference : TDateTime;
begin
  if reference = 0 then
    CurReference := Date()
  else
    CurReference := reference;
  { On début au début de la semaine }
  DateDebut := CurReference - JourDeLaSemaine(CurReference) + 1;
  { On applique le décalage }
  DateDebut := DateDebut - 7 * decalage;
  { On fini en fin de semaine complète }
  DateFin := DateDebut + 6;
end;

function TCalcuVarPre.Str5cDateToStr10cDate(Const RefDate : TDateTime; strDate5c: String): String;
begin
  result := leftStr(strDate5c, 2)+'/'+RightStr(strDate5c,2)+'/'+ IntToStr(YearOf(RefDate));
end;


function TCalcuVarPre.GetTobValeursCompteursSalarie(Salarie: string): Tob;
begin
  result := TobValeursCompteurs.FindFirst(['PYP_SALARIE'], [Salarie], False);
  if not Assigned(result) then
  begin
    result := Tob.Create(TobValeursCompteurs.NomTable, TobValeursCompteurs, -1);
    result.AddChampSupValeur('PYP_SALARIE', Salarie);
  end;
end;

procedure TGestionPresence.GetDatesModeleCycle(var DateDebut, DateFin: TDateTime;
  LaDate: TDateTime; Cycle, TypeDeCycle: String);
var
  tempTob : Tob;
  DiffDates, DureeModele, DureeCycle : Integer;
  Modele : String;
  DateLancementCycle : TDateTime;
begin
  { On vérifie qu'on ai pas déjà calculé l'info par un appel précédent à la fonction GetModeleAndJourneeInCycle}
  if TypeDeCycle = 'MOD' then
  begin
    tempTob := SavTobModeleAndJournee.FindFirst(['TYPEDECYCLE'], [TypeDeCycle], False);
    if Assigned(tempTob) then
    begin
      tempTob := tempTob.FindFirst(['CYCLE'], [Cycle], False);
      if Assigned(tempTob) then
      begin
        tempTob := tempTob.FindFirst(['CHECKEXCEPT'], [False], False);
        if Assigned(tempTob) then
        begin
          tempTob := tempTob.FindFirst(['DATE'], [LaDate], False);
          if Assigned(tempTob) then
          begin
            DateDebut := tempTob.GetDateTime('DEBUTMODELE');
            DateFin   := tempTob.GetDateTime('FINMODELE');
            exit;
          end;
        end;
      end;
    end;
  end;
  Modele := '';
  DateDebut := iDate1900;
  DateFin := iDate2099;
  if TypeDeCycle = 'MOD' then   { Affichage des modèles de cycle ( type défini dans la tablette PGTYPEAFFECTSALARIE) }
  begin
    Modele := Cycle;
    { si Modèle, on cherche la durée du modèle }
    tempTob := TobAssocModele.FindFirst(['PMY_MODELECYCLE'],[Modele],False); { on prends n'importe quel ligne qui correspond au modele }
    if Assigned(tempTob) and (tempTob.GetString('PMY_DATEVALIDITE') <> '') and (tempTob.GetString('PMY_NBJOUR') <> '') then
    begin
      DateLancementCycle := tempTob.GetDateTime('PMY_DATEVALIDITE');
      DiffDates := Trunc(LaDate) - Trunc(DateLancementCycle) +1;
      if DiffDates < 1 then exit; { le modèle n'est pas valide avant sa date de validité }
      DureeModele := tempTob.GetInteger('PMY_NBJOUR');
      DateDebut := DateLancementCycle + DureeModele * Trunc((Trunc(LaDate) - Trunc(DateLancementCycle))/DureeModele);
      DateFin := DateDebut + DureeModele -1;
    end;
  end else begin
    { si cycle, on cherche la durée du cycle }
    tempTob := TobDureeCycle.FindFirst(['PYC_CYCLE'],[Cycle],False);
    if Assigned(tempTob) and (tempTob.GetString('DUREE') <> '') then
    begin
      DateLancementCycle := tempTob.GetDateTime('DATE');
      DiffDates := Trunc(LaDate) - Trunc(DateLancementCycle) +1;
      if DiffDates < 1 then exit; { le cycle n'est pas valide avant sa date de validité }
      DureeCycle := tempTob.GetInteger('DUREE');
      DateDebut := DateLancementCycle + DureeCycle * Trunc((Trunc(LaDate) - Trunc(DateLancementCycle))/DureeCycle);
      DateFin := DateDebut + DureeCycle -1;
    end;
  end;
end;

procedure TCalcuVarPre.FreeMemSalarie(Salarie: String);
var
  TempTob, TempTobPF : Tob;
  indexCompteur, indexPerioDate : Integer;
begin
  TempTob := TobSaveCalcVar.FindFirst(['SALARIE'], [Salarie], False);
  if Assigned(TempTob) then
  begin
    TempTob.ChangeParent(nil,-1);
    FreeAndNil(TempTob);
  end;
  TempTob := TobValeursCompteurs.FindFirst(['PYP_SALARIE'], [Salarie], False);
  if Assigned(TempTob) then
  begin
    { Suppression de la tob des compteurs non mis à jour pour le salarie
     (sinon, insertDB essaye de les insérer à nouveau -> violation de clef) }
    { Parcours des compteurs }
    for indexCompteur := TempTob.Detail.Count -1 downto 0 do
    begin
      { Parcours des periodicités et des dates }
      for indexPerioDate := TempTob.Detail[indexCompteur].detail.count -1 downto 0 do
      begin
        TempTobPF := TempTob.Detail[indexCompteur].Detail[indexPerioDate]; 
        if (not TempTobPF.IsOneModifie) or (TempTobPF.GetDouble('PYP_QUANTITEPRES') = 0) then
        begin
          TempTobPF.ChangeParent(nil,-1);
          FreeAndNil(TempTobPF);
        end;
      end;
    end;
  end;
  TempTob := TobAbsencePresenceSalarie.FindFirst(['PCN_SALARIE'], [Salarie], False);
  if Assigned(TempTob) then
  begin
    TempTob.ChangeParent(nil,-1);
    FreeAndNil(TempTob);
  end;
{  //TobProfilPresenceSalarie  ppz_salarie
  //TobExceptionSalaries PYE_salarie
        }
end;

function TCalcuVarPre.BuildLoadingQuery(DateDebut, DateFin: TDateTime): String;
var
  stSQL2 : String;
  function ConcernePeriode(DateDebut, DateFin : TDateTime) : String;
  begin
    result := '(    (PYP_DATEDEBUTPRES >= "'+USDATETIME(DateDebut)+'" AND PYP_DATEDEBUTPRES <= "'+USDATETIME(DateFin)+'") '
            + '  OR (PYP_DATEFINPRES >= "'  +USDATETIME(DateDebut)+'" AND PYP_DATEFINPRES <= "'  +USDATETIME(DateFin)+'") '
            + '  OR (PYP_DATEDEBUTPRES < "' +USDATETIME(DateDebut)+'" AND PYP_DATEFINPRES > "'  +USDATETIME(DateFin)+'") '
            + ')';
  end;
begin
  result := 'SELECT PYP_SALARIE, PYP_DATEDEBUTPRES, PYP_DATEFINPRES, PYP_COMPTEURPRES, '
         + ' PYP_PERIODICITEPRE, PYP_THEMEPRE, PYP_QUANTITEPRES, PYP_ETATPRES, PYP_PGINDICATPRES, '
         + ' PYP_TYPECALPRES, PYP_DATEDEBUTBUL, PYP_DATEFINBUL, PYP_DATECREATION, PYP_DATEMODIF FROM PRESENCESALARIE '
         ;
  result := result + 'WHERE (';
  { On doit chargé tous les compteurs utilisés par les variables de paie "Date à Date" et "Depuis (hors période en cours)"
    qui ont été calculés depuis le début de l'année }
  if ListeCompteurs412 <> '' then
    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs412+') and '+ConcernePeriode(EncodeDateDay(YearOf(DateDebut),1), DateFin)+')';
  { On doit chargé tous les compteurs utilisés par les variables de paie "Mois en cours" et "Semaines précédentes du mois"
    qui ont été calculés depuis le début du mois }
  if ListeCompteurs813 <> '' then
  begin
    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs813+') and '+ConcernePeriode(EncodeDateMonthWeek(YearOf(DateDebut), MonthOf(DateDebut), 1 ,1), DateFin)+')';
  end;
  { On doit chargé tous les compteurs utilisés par les variables de paie "Semaine en cours"
    qui ont été calculés depuis le début - 1 semaine }
  if ListeCompteurs9 <> '' then
  begin
    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs9+') and '+ConcernePeriode(EncodeDateWeek(YearOf(DateDebut-7), WeekOf(DateDebut-7)), DateFin)+')';
  end;
  { On doit chargé tous les compteurs utilisés par les variables de paie "Mois glissants (hors mois en cours)"
    qui ont été calculés depuis le début -12 mois }
  if ListeCompteurs1014 <> '' then
  begin
    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs1014+') and '+ConcernePeriode(EncodeDateMonthWeek(YearOf(DateDebut)-1, MonthOf(DateDebut), 1 ,1), DateFin)+')';
  end;
  { On doit chargé tous les compteurs utilisés par les variables de paie "Semaine glissante (hors semaine en "
    qui ont été calculés depuis le début -3 mois }
  if ListeCompteurs11 <> '' then
  begin
    if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
    stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in ('+ListeCompteurs11+') and '+ConcernePeriode(EncodeDateMonthWeek(YearOf(IncMonth(DateDebut,-3)), MonthOf(IncMonth(DateDebut,-3)), 1 ,1), DateFin)+')';
  end;
  { DONE 3 -cPrésence -oGGU : Ajouter le chargement des compteurs mensuels susceptibles d'être passés en paie }
  if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
  stSQL2 := stSQL2 + '(PYP_COMPTEURPRES in (select pyr_compteurpres from compteurpresence where pyr_integrepaie = "X") and '+ConcernePeriode(DateDebut, DateFin)+')';
  if stSQL2 <> '' then stSQL2 := stSQL2 + ' or ';
  stSQL2 := stSQL2 + ConcernePeriode(DateDebut, DateFin);
  result := result + stSQL2 + ')';
  AddFiltreSalarie('PYP', result);
end;

end.
