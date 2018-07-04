{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 01/01/1900
Modifié le ... : 22/10/2007
Description .. :
Suite ........ :
Suite ........ : SBO 22/10/2007 : FQ 21681 gestion des erreurs ignorées
Mots clefs ... :
*****************************************************************}
unit ULibAnalytique;

interface

uses
  Classes,
  HCTRLS, // pour le THGrid de GetTOB
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
{$IFNDEF PGIIMMO}
  ULibEcriture, // pour CSetMontants
{$ENDIF}
  SAISUTIL,
  UtilSais,
  UTob,
  HEnt1,
  {$IFNDEF EAGLSERVER}
  hmsgbox, //Hshowmessage  //SG6 21/12/204 FQ 14731
  {$ENDIF}
  ParamSoc, //GetParamSoc
  utilPGI,  // OpenSelect
  SysUtils
  , Ent1
  {$IFDEF MODENT1}
  , CPTypeCons
  , CPProcMetier
  , CPProcGen
  {$ENDIF MODENT1}
  , ed_tools
  ,uEntCommun
  ;

//SG6 16.02.05 Croisaxe
type
  TParamAxe = record
    premier_axe: byte;
    dernier_axe: byte;
  end;
// déplacement de LetBatch
Type R_YV = RECORD
            Axe : String3 ;
            Section : String17 ;
            NumVentil : integer ;
            Pourcentage : Double ;
            Debit       : boolean ;
            TotPivot,TotDevise : double ;
            //SG6 03.03.05 Gestion Mode Croisaxe
            SousPlan1 : string;
            SousPlan2 : string;
            SousPlan3 : string;
            SousPlan4 : string;
            SousPlan5 : string;
           END ;

type
  //Classe pour le passage entre les deux modes analytiques
  TTraitementCroisaxe = class
  private
    {JP 08/11/05 : Pour la gestion du croisement des ventilations types}
    SectAtt1 : string;
    SectAtt2 : string;
    SectAtt3 : string;
    SectAtt4 : string;
    SectAtt5 : string;

    cTobEcr: TOB;
    cTobNew: TOB;
    cRecDevise: RDevise;
    cBooDebit: boolean;
    cBooAnaCroisaxe : boolean;
    cIntNumVentil: integer;
    cArrAxe: array[1..MaxAxe] of boolean;
    cIntPremAxe: integer;
    cImport : boolean; //Imports ?
    cSoldeSection : TOB;
    cSoldeJournal : TOB;
    function GetPourCent(Section: TOB): Double;
    function CalculMontant(NewSec: TOB; vDevise: boolean): Double;
    function NewAna(S1 : TOB; S2 : TOB = nil ; S3: TOB = nil; S4: TOB = nil; S5: TOB = nil): Double;
    procedure VerifEquil(Mnt, MntTmp, MntDev, MntDevTmp, Pct, PctTmp: Double);
    procedure VerifEquilCroisaxe;
    function GetTobAxe(Ind: Integer): TOB;
    procedure InitClass(vTobEcr: TOB);
    procedure MajSoldeSectionDelete(vAxe : integer);
    function RechercheJournalOD(vIntAxe : integer):string;
    procedure AjoutSoldeSection(vTob : TOB);
    procedure AjoutSoldeJournal(vTob : TOB; vDblDeb, vDblCre : double);

    {JP 08/11/05 : Pour la gestion du croisement des ventilations types}
    procedure SetSectAtt;
    {JP 09/11/05 : Si le total des pourcentages <> 100, on fait une régularisation}
    procedure GereArrondi(Nat, Cpte : string; var TobN : TOB);
  public

    {JP 02/08/07 : Décroisement des axes : on mémorise les axes qui étaient croisés}
    procedure MajAxeTraitement(TabAxe : array of Boolean);

    procedure MajSoldeSection;
    procedure DeleteInsertAna;
    procedure SetTobEcr(vTobEcr: TOB);
    function PassageVentAnaClassiqueCroisaxe: TOB;
    function PassageVentCroisaxeAnaClassique: TOB;
    function PassageAnaPurClassiqueCroisaxe: TOB;
    function PassageAnaPurCroisaxeClassique: TOB;
    {JP 07/11/05 : On croise aussi le paramétrage, à tout le moins les Ventilations Types}
    procedure CroiseVentilTypeEtAutres;
    {JP 07/11/05 : Suppression du paramétrage avant insertion des ventilations croisées}
    function SupprimeVentiType(Tout : Boolean) : Boolean;
    {JP 07/11/05 : Calcul de la nouvelle ventilation pour une nature et un compte donnés}
    procedure GetLesPourcentage(Cpte, Nat : string; var TobN, Tob1, Tob2, Tob3, Tob4, Tob5 : TOB);

    constructor Create(vA1, vA2, vA3, vA4, vA5, vImport: boolean; vTobEcr: TOB = nil);overload;
    constructor Create(vImport: boolean; vTobEcr: TOB = nil); overload;
    destructor Destroy; override;
  end;

type
  TRestrictionAnalytique = class
  private
    FCompteGene: String;
    FModele:     array[1..MaxAxe] of String;
    FExclu:      array[1..MaxAxe] of Boolean;
    FChampAna:   array[1..MaxAxe] of String;
    FAxeCroise: Boolean;
    procedure SetAxeCroise(const Value: Boolean);
  protected
    procedure InitCompteGene(CompteGene: String);
    function  GetFirstAxe(Axe: String): Integer;
    function  GetClause(Axe, TableAna: String; CompteAna: array of String): String;
    function GetClauseAttente(iAxe: Integer): String;
  public
    constructor Create;
    destructor  Destroy; override;

    {Retourne une clause SQL permettant de connaitre les comptes analytiques répondant à la déf du modèle}
    function GetClauseCompteAutorise(
                 CompteGene: String;  {Nom du compte géné permettant de retrouver le modèle associé}
                 Axe: String;         {Axe analytique des comptes à vérifier}
                 TableAna: String;    {Nom de la table ayant le champ à vérifier}
                 CompteAna: array of String  {Tableau contenant les valeurs comptes des autres axes. Que pour les axes croisés et la table SECTION}
                 ): String;

    {Retourne la jointure SQL permettant de connaitre les comptes analytiques contraire à la déf du modèle}
    function GetFromCompteInterdit(
                 Axe: String;         {Axe analytique des comptes à vérifier}
                 ChampCompteGene: String;    {Nom du champ Compte Géné}
                 ChampCompteAna: String      {Nom du champ Compte Analytique}
                 ): String;

    {Retourne une clause SQL permettant de connaitre les comptes analytiques contraire à la déf du modèle}
    function GetClauseCompteInterdit(
                 Axe: String;         {Axe analytique des comptes à vérifier}
                 TableAna: String     {Nom de la table ayant le champ à vérifier}
                 ): String; overload;

    {Retourne une clause SQL permettant de connaitre les comptes analytiques contraire à la déf du modèle}
    function GetClauseModeleCompteInterdit(
                 Modele: String;      {Nom du modèle de restriction à vérifier}
                 Axe: String;         {Axe analytique des comptes à vérifier}
                 TableAna: String     {Nom de la table ayant le champ à vérifier}
                 ): String;

    {Retourne True si le modèle de ventil est utilisable par un compte géné}
    function IsCompteGeneAutorise(
                 CompteGene: String;  {Nom du compte géné permettant de retrouver le modèle associé}
                 Axe: String;         {Axe analytique du modèle}
                 Nature: String;      {Nature du modèle}
                 Modele: String       {Nom du modèle de ventilation à vérifier}
                 ): Boolean;

    {Vérifie que le modèle de ventil n'a pas de compte interdit}
    function VerifModelVentil(
                 CompteGene: String;    {Nom du compte géné permettant de retrouver le modèle associé}
                 Q: TQuery;             {Query contenant les comptes du modèle de ventil sans les comptes ana exclus par la restriction}
                 Table,                 {Table contenant le modèle de ventil}
                 Axe: String;           {Axe analytique du modèle}
                 QrySansFiltre: String;
                 ForceMsgErr: Boolean = False
                 ): Boolean; overload;

    {Vérifie que le modèle de ventil n'a pas de compte interdit: uniquement pour CPVENTILTYPECROIS_TOF}
    function VerifModelVentil(CompteGene: String; Q, QAll: TOB; Table, Axe: String; ForceMsgErr: Boolean = False): Boolean; overload;

    {Renvoie True si le compte est associé à un modèle qui exclu certaines sections}
    function IsModelExclu(
                 CompteGene: String;    {Nom du compte géné permettant de retrouver le modèle associé}
                 Axe:        String     {Axe analytique du modèle}
                 ): Boolean;

    property AxeCroise: Boolean read FAxeCroise write SetAxeCroise;  {Pour les budgets, pas d'axes croisés}

  end;

  //====================================================
  //===== NOUVELLES FONCTIONS UTILITAIRES POUR TOB =====
  //====================================================
function GetMontant(Obj: TOB): Double; // Retourne le montant (crédit ou débit)
function GetMontantDev(Obj: TOB): Double; // Retourne le montant devise (crédit ou débit)
function GetSens(Obj: TOB): Integer; // Retourne le sens de l'écriture (debit=1; crédit=2)
function VentilationExiste(Obj: TOB): Boolean; // Retourne vrai si au moins un des axes est ventilé
function GetCreditSaisi(Obj: TOB): double; // Retourne le montant crédit saisie en fonction du mode opposé
function GetDebitSaisi(Obj: TOB): double; // Retourne le montant débit saisie en fonction du mode opposé
function GetTobPrefixe(Obj: TOB): string; // Retourne le préfixe de la table associée à la tob

// Procédure de création et d'initialisation des TOB analytique
procedure ChargeAnalytique( vTobPiece : TOB ; vDossier : String = '' ); // Chargement de l'analytique dans une tob de type "Pièce"
procedure CPutDefautAna(vTOBLigneAna: TOB); // Initialise une TOB type Analytique
procedure CAddChampSupAna(vTOBLigneAna: TOB); // Ajoute les champs sup puor la gestion analytique
function CGetNewTOBAna(NumAxe: Integer; Parent: TOB = nil): TOB; // Création d'une TOB Analytique + Initialisation

// Les fonctions suivantes prennent en paramètre des objets (TOB) basé sur l'écriture ventilée
// Structure : TOB("ECRITURE") --> TOB("AX") --> TOB("ANALYTIQ")
function GetTotalVentil(ObjEcr: TOB; NumAxe: Integer = 0): double; // Retourne le cumul des ventilations pour l'axe NumAxe
function GetTotalVentilDev(ObjEcr: TOB; NumAxe: Integer = 0): double; // Retourne le cumul des ventilations pour l'axe NumAxe
function GetSensVentil(ObjEcr: TOB; NumAxe: Integer = 0): double; // Retourne le sens de la ventilation (debit=1; crédit=2) pour l'axe NumAxe
function GetVentilGeneral(ObjEcr: TOB; NumAxe: Integer = 0): string; // Retourne le compte géné renseigné dans les ventilations
function GetModeConfVentil(ObjEcr: TOB; NumAxe: Integer = 0): string; // Retourne le mode de confidentialité des ventilations de l'écriture le plus élevé
procedure InitCommunObjAnalNEW(ObjEcr: TOB; ObjAna: TOB); // Init d'une TOB Analytique avec une TOB Ecriture
procedure AlloueAxe(ObjEcr: TOB); // Aloue les TOB Axe à une TOB écriture

// Procédure utilisé dans la gestion des interfaces
function GetTOB(GS: THGrid; Lig: integer): TOB; // Retourne la TOB associé à la ligne Lig de la grille GS
procedure CChargeAna(vTOB: TOB);
function CGetCleAna(vTOB: TOB): string;
function CGetNumAxe(vAxe: string): Integer;

// Fonctions de validation de l'analytique
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
function  CIsValidSection    ( vTobLigne : Tob ; vInfo : TInfoEcriture ) : integer ;
function  CIsValidMontantAna ( vTobLigne : Tob ; vInfo : TInfoEcriture ) : integer ;
function  CIsValidLigneAna   ( vTobLigne : Tob ; vInfo : TInfoEcriture ) : TRecError ;
function  CIsValidAxeBudget  ( vAxe : integer ; vTobAxe : Tob ; vInfo : TInfoEcriture ) : integer ;
function  CIsValidAxe        ( vAxe : integer ; vTobAxe : Tob ; vInfo : TInfoEcriture ; vBoVerifQte : Boolean = False ; vBoVerifBudget : Boolean = False ) : TRecError ;
function  CIsValidVentil     ( vTobEcr   : Tob ; vInfo : TInfoEcriture ; vBoVerifQte : Boolean = False ; vBoVerifBudget : Boolean = False ) : TRecError ;
function  CAOuvrirVentil     ( vTobEcr   : Tob ; vTobScenario : Tob ; vInfo : TInfoEcriture ; var vNumAxe : integer ; vBoParam : Boolean = False ) : Boolean ;
procedure CGereFenAna        ( vStGen, vStNatJal : string ; vBoUnSeulAxe : boolean ; var vBoOuvreAna , vBoGestionTva , vBoVentilTVA : boolean ) ;
procedure CVentilerTOB       ( vTOBEcr : TOB ; vInfo : TInfoEcriture ; QueAttente : boolean = false ; vCodeGuide : string = '' ) ;
{$ENDIF}
{$ENDIF}

// SBO 26/03/2007 DEV 5942 : Affectation auto des paramètres de ventilation
function InitialiseVentilDefaut ( Compte : string ; SectAtt : Array of Boolean ; var AxeVent : Array of Boolean ) : boolean ; // FQ 21259 SBO 27/08/2007
function  ChercheCompteDansRub(CptRub,ComptGene : String) : Boolean ;



//=========================================================================
//===== NOUVELLES PROCEDURES DE PRORATISATION DE LA TVA AVEC DES TOBS =====
//=========================================================================
procedure RecalculProrataAnalNEW(Pf: string; ObjEcr: TOB; NumA: integer; DEV: RDevise);

//=========================================================================================================
//===== FONCTION PERMETTANT DE VERIFIER LES QUANTITES SAISIES ENTRE LIGNE D ECRITURE ET ANALYTIQUES =======
//===== SG6 21/12/2004 FQ 14731                                                                     =======
//=========================================================================================================
function CheckQuantite( ObjEcr: TOB ; vBoParle : Boolean = True ) : Boolean ;

//=========================================================================================================
//===== FONCTIONS ET PROCEDURE POUR LE MODE ANALYTIQUE CROISAXE                                     =======
//===== SG6 16.02.05                                                                                =======
//=========================================================================================================
//Recherche le premier ventilable
function RecherchePremDerAxeVentil: TParamAxe;
//Creer la ventilation de la la société lors du passage en mode croisaxe d'une sociéte

{Retourne le premier  axe croisé ou Axe}
function GetFirstAxeAna(Axe: String): Integer;
{JP 24/10/07 : Retourne les infos d'un axe en tenant compte du partage et du dossier}
function GetInfosAxe(Axe : Integer; aDossier : string = '') : TInfoCpta;

Procedure CReporteVentil ( vTobEcr : TOB ; vChamps : String ; vValeur : Variant ; vBoForce : Boolean = False ) ;
Procedure CSynchroVentil ( vTobEcr : TOB ) ;

procedure MajVentilation(LeCpte: String ; VentilAxe : array of string ; PreVentil : Boolean; aDossier : string = '');
{ déplacement del LetBatch}
Procedure EcrToAna ( OEcr,OAna : TOBM ) ;
Procedure FiniMontantsAna ( OAna : TOBM ; RY : R_YV ; DEV : RDEVISE ; UnSeul : boolean; aDossier : string = '') ;
{ déplacement de aglstdcompta LetBatch}
procedure MajGVentil(OKgene : Boolean; LeCpte : string);

{$IFNDEF TRESO}
{$ENDIF}
implementation

const MaxEcriture : integer = 9999;

Const
  // Champs à synchroniser Ecr -> Ventil
  _InMaxChampsSynchVentil = 35 ;      {FP 02/05/2006}
  _recChampsSynchVentil : array[1.._InMaxChampsSynchVentil] of string =
      ('E_GENERAL'          ,'E_AUXILIAIRE'        ,'E_REFINTERNE'        ,'E_LIBELLE'
      , 'E_NATUREPIECE'     ,'E_QUALIFPIECE'       ,'E_QUALIFQTE2'        ,'E_QTE1'
      , 'E_QTE2'            ,'E_ETABLISSEMENT'     ,'E_JOURNAL'           ,'E_CONTREPARTIEGEN'
      , 'E_CONTREPARTIEAUX' ,'E_QUALIFQTE1'        ,'E_NUMEROPIECE'       ,'E_NUMLIGNE'
      , 'E_DATECOMPTABLE'   ,'E_ECRANOUVEAU'       ,'E_EXERCICE'          ,'E_NUMEROPIECE'
      , 'E_VALIDE'          ,'E_UTILISATEUR'       ,'E_DATECREATION'      ,'E_DATEMODIF'
      , 'E_SOCIETE'         ,'E_DEVISE'            ,'E_TYPEMVT'           ,'E_CONFIDENTIEL'
      , 'E_CREERPAR'        ,'E_PERIODE'           ,'E_SEMAINE'
      , 'E_DEBIT'           ,'E_CREDIT'            ,'E_DEBITDEV'          ,'E_CREDITDEV'   {FP 02/05/2006}
      ) ;


//====================================================
//===== NOUVELLES FONCTIONS UTILITAIRES POUR TOB =====
//====================================================

function CGetNumAxe(vAxe: string): Integer;
begin
  Result := StrToInt(Copy(vAxe, 2, 1));
end;

function CGetCleAna(vTOB: TOB): string;
begin
  result := 'Y_JOURNAL="' + vTOB.GetValue('E_JOURNAL') + '" AND ' +
    'Y_EXERCICE="' + vTOB.GetValue('E_EXERCICE') + '" AND ' +
    'Y_DATECOMPTABLE ="' + UsDateTime(vTOB.GetValue('E_DATECOMPTABLE')) + '" AND ' +
    'Y_NUMEROPIECE=' + IntToStr(vTOB.GetValue('E_NUMEROPIECE')) + ' AND ' +
    'Y_NUMLIGNE=' + IntToStr(vTOB.GetValue('E_NUMLIGNE')) + ' AND ' +
    'Y_QUALIFPIECE="' + vTOB.GetValue('E_QUALIFPIECE') + '" AND ' +
    'Y_CONTROLEUR="' + vTOB.GetValue('E_CONTROLEUR') + '" ';
end;

procedure CChargeAna(vTOB: TOB);
var
  lQ: TQuery;
  lSt: string;
  lInIndexA: integer;
begin

  lQ := nil;

  if vTOB.GetString('E_ANA') <> 'X' then exit ;

  AlloueAxe(vTOB);

  // SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
  lSt := 'SELECT ' + GetSelectAll('Y', True) + ', Y_BLOCNOTE FROM ANALYTIQ WHERE Y_JOURNAL="' + vTOB.GetValue('E_JOURNAL') + '" ' +
    'AND Y_EXERCICE="' + vTOB.GetValue('E_EXERCICE') + '" ' +
    'AND Y_DATECOMPTABLE="' + usDateTime(vTOB.GetValue('E_DATECOMPTABLE')) + '" ' +
    'AND Y_EXERCICE="' + vTOB.GetValue('E_EXERCICE') + '" ' +
    'AND Y_NUMEROPIECE=' + intToStr(vTOB.GetValue('E_NUMEROPIECE')) + ' ' +
    'AND Y_NUMLIGNE=' + intToStr(vTOB.GetValue('E_NUMLIGNE')) + ' ' +
    'AND Y_QUALIFPIECE="' + vTOB.GetValue('E_QUALIFPIECE') + '" ';

  try
    lQ := OpenSql(lSt, true,-1,'',true);

    while not lQ.EOF do
    begin
      lInIndexA := CGetNumAxe(lQ.FindField('Y_AXE').asString);
      TOB.CreateDB('ANALYTIQ', vTOB.Detail[lInIndexA - 1], -1, lQ);
      lQ.Next;
    end; //while


  finally
    Ferme(lQ);
  end; // try

end;

function GetTobPrefixe(Obj: TOB): string;
begin
  Result := TableToPrefixe(Obj.NomTable);
end;

function GetMontantDev(Obj: TOB): Double;
var
  StPref: string;
begin
  Result := 0;
  if Obj = nil then Exit;
  if Obj.NomTable = 'BUDECR' then Exit; // Attention ! Montant DEV non géré sur BUDECR
  {	if Obj.NomTable = 'ECRITURE' then StPref := 'E'
    else if Obj.NomTable = 'ANALYTIQ' then StPref := 'Y'
      else Exit;}
  stPref := GetTobPrefixe(Obj);
  if stPref = '' then Exit;

  Result := Obj.GetValue(StPref + '_DEBITDEV') + Obj.GetValue(StPref + '_CREDITDEV');
end;

function GetMontant(Obj: TOB): Double;
var
  StPref: string;
begin
  Result := 0;
  if Obj = nil then Exit;
  {	if Obj.NomTable = 'ECRITURE' then StPref := 'E'
    else if Obj.NomTable = 'ANALYTIQ' then StPref := 'Y'
     else if Obj.NomTable = 'BUDECR' then StPref := 'BE'
       else Exit;}
  stPref := GetTobPrefixe(Obj);
  if stPref = '' then Exit;
  Result := Obj.GetValue(StPref + '_DEBIT') + Obj.GetValue(StPref + '_CREDIT');
end;

function GetSens(Obj: TOB): Integer;
var
  StPref: string;
begin
  Result := 0;
  if Obj = nil then Exit;
  {	if Obj.NomTable = 'ECRITURE' then StPref := 'E'
    else if Obj.NomTable = 'ANALYTIQ' then StPref := 'Y'
     else if Obj.NomTable = 'BUDECR' then StPref := 'BE'
       else Exit;}
  stPref := GetTobPrefixe(Obj);
  if stPref = '' then Exit;
  {$IFDEF COMPTA} //SG6 02/12/2004 Différence de fonctionnement entre une V4 et V6
  if (Obj.GetValue(StPref + '_DEBIT')) <> 0
    then Result := 1
  else Result := 2;
  {$ELSE}
  if (Obj.GetValue(StPref + '_DEBIT')) > 0
    then Result := 1
  else Result := 2;
  {$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 29/10/2002
Modifié le ... :   /  /
Description .. : Retourne True si l'écriture passée en paramètre est ventilée
Suite ........ : sur au moins un axe.
Suite ........ : Structure de la TOB attendue :
Suite ........ :
Suite ........ :  TOB("ECRITURE") --> TOB("AX") --> TOB("ANALYTIQ")
Mots clefs ... :
*****************************************************************}
function VentilationExiste(Obj: TOB): Boolean;
var
  i: Integer;
begin
  Result := False;
  if Obj = nil then Exit;
  for i := 0 to Obj.Detail.Count - 1 do
    if Obj.Detail[i].Detail.Count > 0 then
    begin
      Result := True;
      Exit;
    end;
end;

function GetCreditSaisi(Obj: TOB): double;
var
  StPref: string;
begin
  Result := 0;
  if Obj = nil then Exit;
  {	if Obj.NomTable = 'ECRITURE' then StPref := 'E'
    else if Obj.NomTable = 'ANALYTIQ' then StPref := 'Y'
     else if Obj.NomTable = 'BUDECR' then StPref := 'BE'
       else Exit;}
  stPref := GetTobPrefixe(Obj);
  if stPref = '' then Exit;

  if Obj.NomTable = 'BUDECR'
    then Result := Obj.GetValue(StPref + '_CREDIT')
  else Result := Obj.GetValue(StPref + '_CREDITDEV');
end;

function GetDebitSaisi(Obj: TOB): double;
var
  StPref: string;
begin
  Result := 0;
  if Obj = nil then Exit;
  {	if Obj.NomTable = 'ECRITURE' then StPref := 'E'
    else if Obj.NomTable = 'ANALYTIQ' then StPref := 'Y'
     else if Obj.NomTable = 'BUDECR' then StPref := 'BE'
       else Exit;}
  stPref := GetTobPrefixe(Obj);
  if stPref = '' then Exit;
  if Obj.NomTable = 'BUDECR'
    then Result := Obj.GetValue(StPref + '_DEBIT')
  else Result := Obj.GetValue(StPref + '_DEBITDEV');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/10/2002
Modifié le ... : 30/10/2002
Description .. : Fonction d'initialisation d'une nouvelle ligne d'analytique
Mots clefs ... :
*****************************************************************}
procedure CPutDefautAna(vTOBLigneAna: TOB);
begin
  vTOBLigneAna.putValue('Y_EXERCICE', GetEntree.Code);
  vTOBLigneAna.putValue('Y_NATUREPIECE', 'OD');
  vTOBLigneAna.putValue('Y_QUALIFPIECE', 'N');
  vTOBLigneAna.putValue('Y_ETAT', '0000000000');
  vTOBLigneAna.putValue('Y_UTILISATEUR', V_PGI.User);
  vTOBLigneAna.putValue('Y_DATECREATION', Date);
  vTOBLigneAna.putValue('Y_DATEMODIF', NowH);
  vTOBLigneAna.putValue('Y_SOCIETE', V_PGI.CodeSociete);
  vTOBLigneAna.putValue('Y_ETABLISSEMENT', GetParamsoc('SO_ETABLISDEFAUT'));
  vTOBLigneAna.putValue('Y_DEVISE', V_PGI.DevisePivot);
  vTOBLigneAna.putValue('Y_TAUXDEV', 1);
  vTOBLigneAna.putValue('Y_DATETAUXDEV', V_PGI.DateEntree);
  vTOBLigneAna.putValue('Y_QUALIFQTE1', '...');
  vTOBLigneAna.putValue('Y_QUALIFQTE2', '...');
  vTOBLigneAna.putValue('Y_QUALIFECRQTE1', '...');
  vTOBLigneAna.putValue('Y_QUALIFECRQTE2', '...');
  vTOBLigneAna.putValue('Y_ECRANOUVEAU', 'N');
  vTOBLigneAna.putValue('Y_CREERPAR', 'SAI');
  vTOBLigneAna.putValue('Y_EXPORTE', '---');
  vTOBLigneAna.putValue('Y_CONFIDENTIEL', '0');
  // V9 CEGID
  vTOBLigneAna.PutValue('Y_DATPER',iDate1900) ;
  vTOBLigneAna.PutValue('Y_ENTITY',0) ;
  vTOBLigneAna.PutValue('Y_REFGUID','') ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/10/2002
Modifié le ... :   /  /
Description .. : Ajoute les champs sup puor la gestion analytique
Mots clefs ... :
*****************************************************************}
procedure CAddChampSupAna(vTOBLigneAna: TOB);
begin
  vTOBLigneAna.addChampSup('OLDDEBIT', FALSE);
  vTOBLigneAna.addChampSup('OLDCREDIT', FALSE);
  vTOBLigneAna.addChampSup('BADROW', FALSE);
  vTOBLigneAna.putValue('OLDDEBIT', vTOBLigneAna.GetValue('Y_DEBIT'));
  vTOBLigneAna.putValue('OLDCREDIT', vTOBLigneAna.GetValue('Y_CREDIT'));
  vTOBLigneAna.putValue('BADROW', '-');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 30/10/2002
Modifié le ... :   /  /
Description .. : Création d'une TOB Analytique + Initialisation
Mots clefs ... :
*****************************************************************}
function CGetNewTOBAna(NumAxe: Integer; Parent: TOB): TOB;
var
  vTOBLigneAna: TOB;
begin
  vTOBLigneAna := TOB.Create('ANALYTIQ', Parent, -1);
  vTobLigneAna.InitValeurs ;
  CAddChampSupAna(vTOBLigneAna);
  CPutDefautAna(vTOBLigneAna);
  if NumAxe > 0 then vTOBLigneAna.putValue('Y_AXE', 'A' + Chr(48 + NumAxe));
  Result := vTOBLigneAna;
end;

function GetTotalVentil(ObjEcr: TOB; NumAxe: Integer): double;
var
  i, j: Integer;
begin
  Result := 0;
  if ObjEcr = nil then Exit;
  if ObjEcr.Detail.Count = 0 then Exit;
  for i := 0 to ObjEcr.Detail.Count - 1 do
    if ((NumAxe = 0) or (NumAxe = i)) and (ObjEcr.Detail[i].Detail.Count > 0) then
    begin
      for j := 0 to ObjEcr.Detail[i].Detail.Count - 1 do
        Result := Result + GetMontant(ObjEcr.Detail[i].Detail[j]);
      Exit;
    end;
end;

function GetTotalVentilDev(ObjEcr: TOB; NumAxe: Integer): double;
var
  i, j: Integer;
begin
  Result := 0;
  if ObjEcr = nil then Exit;
  if ObjEcr.Detail.Count = 0 then Exit;
  for i := 0 to ObjEcr.Detail.Count - 1 do
    if ((NumAxe = 0) or (NumAxe = i)) and (ObjEcr.Detail[i].Detail.Count > 0) then
    begin
      for j := 0 to ObjEcr.Detail[i].Detail.Count - 1 do
        Result := Result + GetMontantDev(ObjEcr.Detail[i].Detail[j]);
      Exit;
    end;
end;

function GetSensVentil(ObjEcr: TOB; NumAxe: Integer): double;
var
  i, j: Integer;
begin
  Result := 0;
  if ObjEcr = nil then Exit;
  if ObjEcr.Detail.Count = 0 then Exit;
  for i := 0 to ObjEcr.Detail.Count - 1 do
    if ((NumAxe = 0) or (NumAxe = i)) and (ObjEcr.Detail[i].Detail.Count > 0) then
    begin
      for j := 0 to ObjEcr.Detail[i].Detail.Count - 1 do
        Result := GetSens(ObjEcr.Detail[i].Detail[j]);
      Exit;
    end;
end;

function GetVentilGeneral(ObjEcr: TOB; NumAxe: Integer): string;
var
  i: Integer;
begin
  Result := '';
  if ObjEcr = nil then Exit;
  if ObjEcr.Detail.Count = 0 then Exit;
  for i := 0 to ObjEcr.Detail.Count - 1 do
    if ((NumAxe = 0) or (NumAxe = i)) and (ObjEcr.Detail[i].Detail.Count > 0) then
    begin
      Result := ObjEcr.Detail[i].Detail[0].GetValue('Y_GENERAL');
      Exit;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 26/04/2007
Modifié le ... :   /  /    
Description .. : - LG - 26/04/2007 - FB 20105 - on affecte le champs 
Suite ........ : Y_AUXILIAIRE
Mots clefs ... : 
*****************************************************************}
procedure InitCommunObjAnalNEW(ObjEcr: TOB; ObjAna: TOB);
var
  PfA, PfE: string[2];
begin
  if (ObjEcr = nil) or (ObjAna = nil) then Exit;
  PfA := 'Y';
  if ObjEcr.NomTable = 'ECRITURE' then
  begin
    PfE := 'E';
    ObjAna.PutValue(PfA + '_GENERAL', ObjEcr.GetValue(PfE + '_GENERAL'));
    ObjAna.PutValue(PfA + '_AUXILIAIRE', ObjEcr.GetValue(PfE + '_AUXILIAIRE'));
    ObjAna.PutValue(PfA + '_REFEXTERNE', ObjEcr.GetValue(PfE + '_REFEXTERNE'));
    ObjAna.PutValue(PfA + '_DATEREFEXTERNE', ObjEcr.GetValue(PfE + '_DATEREFEXTERNE'));
    ObjAna.PutValue(PfA + '_AFFAIRE', ObjEcr.GetValue(PfE + '_AFFAIRE'));
    {$IFNDEF SPEC302}
    ObjAna.PutValue(PfA + '_PERIODE', GetPeriode(ObjEcr.GetValue(PfE + '_DATECOMPTABLE')));
    ObjAna.PutValue(PfA + '_SEMAINE', NumSemaine(ObjEcr.GetValue(PfE + '_DATECOMPTABLE')));
    {$ENDIF}
  end
  else if ObjEcr.NomTable = 'BUDECR' then PfE := 'BE'
  else Exit;

  ObjAna.PutValue(PfA + '_EXERCICE', ObjEcr.GetValue(PfE + '_EXERCICE'));
  ObjAna.PutValue(PfA + '_DATECOMPTABLE', ObjEcr.GetValue(PfE + '_DATECOMPTABLE'));
  ObjAna.PutValue(PfA + '_NUMEROPIECE', ObjEcr.GetValue(PfE + '_NUMEROPIECE'));
  ObjAna.PutValue(PfA + '_NUMLIGNE', ObjEcr.GetValue(PfE + '_NUMLIGNE'));
  ObjAna.PutValue(PfA + '_REFINTERNE', ObjEcr.GetValue(PfE + '_REFINTERNE'));
  ObjAna.PutValue(PfA + '_LIBELLE', ObjEcr.GetValue(PfE + '_LIBELLE'));
  ObjAna.PutValue(PfA + '_NATUREPIECE', ObjEcr.GetValue(PfE + '_NATUREPIECE'));
  ObjAna.PutValue(PfA + '_UTILISATEUR', ObjEcr.GetValue(PfE + '_UTILISATEUR'));
  ObjAna.PutValue(PfA + '_CONTROLEUR', ObjEcr.GetValue(PfE + '_CONTROLEUR'));
  ObjAna.PutValue(PfA + '_SOCIETE', ObjEcr.GetValue(PfE + '_SOCIETE'));
  ObjAna.PutValue(PfA + '_ETABLISSEMENT', ObjEcr.GetValue(PfE + '_ETABLISSEMENT'));

  ObjAna.PutValue(PfA + '_TAUXDEV', ObjEcr.GetValue(PfE + '_TAUXDEV'));
  ObjAna.PutValue(PfA + '_DATETAUXDEV', ObjEcr.GetValue(PfE + '_DATETAUXDEV'));
  ObjAna.PutValue(PfA + '_DEVISE', ObjEcr.GetValue(PfE + '_DEVISE'));
  ObjAna.PutValue(PfA + '_JOURNAL', ObjEcr.GetValue(PfE + '_JOURNAL'));

  ObjAna.PutValue(PfA + '_TOTALECRITURE', GetMontant(ObjEcr));
  ObjAna.PutValue(PfA + '_TOTALDEVISE', GetMontantDev(ObjEcr));
  ObjAna.PutValue(PfA + '_QUALIFPIECE', ObjEcr.GetValue(PfE + '_QUALIFPIECE'));

  {infos supplémentaires}
  ObjAna.PutValue(PfA + '_QUALIFECRQTE1', ObjEcr.GetValue(PfE + '_QUALIFQTE1'));
  ObjAna.PutValue(PfA + '_QUALIFECRQTE2', ObjEcr.GetValue(PfE + '_QUALIFQTE2'));
  ObjAna.PutValue(PfA + '_QUALIFQTE1', ObjEcr.GetValue(PfE + '_QUALIFQTE1'));
  ObjAna.PutValue(PfA + '_QUALIFQTE2', ObjEcr.GetValue(PfE + '_QUALIFQTE2'));
  ObjAna.PutValue(PfA + '_TOTALQTE1', ObjEcr.GetValue(PfE + '_QTE1'));
  ObjAna.PutValue(PfA + '_TOTALQTE2', ObjEcr.GetValue(PfE + '_QTE2'));
  ObjAna.PutValue(PfA + '_REFLIBRE', ObjEcr.GetValue(PfE + '_REFLIBRE'));

end;

procedure AlloueAxe(ObjEcr: TOB);
var
  NumAxe: Integer;
  lTobTmp: TOB;
begin
  if ObjEcr = nil then Exit;
  if ObjEcr.Detail.Count < MaxAxe then
  begin
    ObjEcr.ClearDetail;
    for NumAxe := 1 to MaxAxe do
    begin
      lTobTmp := TOB.Create('$AXE', ObjEcr, -1);
      //SG6 29.03.05 Rajoute d'un champ sup Y_AXE
      lTobTmp.AddChampSupValeur('Y_AXE', 'A' + IntTostr(NumAxe));
    end;

  end;
end;

function GetModeConfVentil(ObjEcr: TOB; NumAxe: Integer): string;
var
  i, j: Integer;
  stPfA: string[2];
begin
  Result := '0';
  if ObjEcr = nil then Exit;
  if ObjEcr.Detail.Count = 0 then Exit;
  stPfA := 'Y';
  for i := 0 to ObjEcr.Detail.Count - 1 do
    if (NumAxe = 0) or (NumAxe = i) then
      for j := 0 to ObjEcr.Detail[i].Detail.Count - 1 do
        if ValeurI(Result) < ValeurI(ObjEcr.Detail[i].Detail[j].GetValue(stPfA + '_CONF'))
          then Result := ObjEcr.Detail[i].Detail[j].GetValue(stPfA + '_CONF');
end;

function GetTOB(GS: THGrid; Lig: integer): TOB;
begin
  case GS.TypeSais of
    tsAnal: Result := TOB(GS.Objects[AN_NumL, Lig]);
    tsLettrage: Result := TOB(GS.Objects[GS.ColCount - 1, Lig]);
    tsPointage: Result := TOB(GS.Objects[0, Lig]);
  else Result := TOB(GS.Objects[SA_Exo, Lig]);
  end;
end;

//==============================================================================
//===== NOUVELLES PROCEDURES DE PRORATISATION DE LA TVA AVEC DES TOBS ==========
//==============================================================================

procedure CalculTotalAnalNEW(Pf: string; ObjEcr: TOB; NumAxe: Integer; Lig: Integer; var MtP, MtD, MtE, TauxAN: Double);
var
  i: Integer;
  ObjAna: TOB;
begin
  Mtp := 0;
  MtD := 0;
  MtE := 0;
  TauxAn := 0;
  for i := 0 to ObjEcr.Detail[NumAxe - 1].Detail.Count - 1 do
    if i <> Lig then
    begin
      ObjAna := ObjEcr.Detail[NumAxe - 1].Detail[i];
      TauxAn := TauxAn + ObjAna.GetValue(Pf + '_POURCENTAGE');
      if GetSens(ObjAna) = 1 then
      begin
        MtP := MtP + ObjAna.GetValue(Pf + '_DEBIT');
        MtD := MtD + ObjAna.GetValue(Pf + '_DEBITDEV');
      end
      else
      begin
        MtP := MtP + ObjAna.GetValue(Pf + '_CREDIT');
        MtD := MtD + ObjAna.GetValue(Pf + '_CREDITDEV');
      end;
    end;
end;

procedure InverseSensAnaNew(ObjEcr: TOB; NumAxe: Integer);
var
  ObjAna: TOB;
  D, C: Double;
  NumV: integer;
begin
  for NumV := 0 to ObjEcr.Detail[NumAxe - 1].Detail.Count - 1 do
  begin
    ObjAna := ObjEcr.Detail[NumAxe - 1].Detail[NumV];
    // Montants
    D := ObjAna.GetValue('Y_DEBIT');
    C := ObjAna.GetValue('Y_CREDIT');
    ObjAna.PutValue('Y_DEBIT', C);
    ObjAna.PutValue('Y_CREDIT', D);
    // Montant devise
    D := ObjAna.GetValue('Y_DEBITDEV');
    C := ObjAna.GetValue('Y_CREDITDEV');
    ObjAna.PutValue('Y_DEBITDEV', C);
    ObjAna.PutValue('Y_CREDITDEV', D);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : - LG - 29/08/2007 - FB 21310 - correction du chg de sens 
Suite ........ : d'un ligne d'ecriture
Mots clefs ... : 
*****************************************************************}
procedure Prorate1AnalNEW(Pf: string; ObjEcr: TOB; NumAxe: Integer; DEV: RDEVISE ; var MtP, MtD, MtE, TauxAN: Double);
var
  LeTaux: Double;
  MtEcrP, MtEcrD : Double;
  TotP, TotD : Double;
  FDebit, FCredit, FDebitDev, FCreditDev: Double;
  ObjAna: TOB;
  NumV, Sens: Integer;
begin
  // Test ventil et montant à proratiser non nul
  if ObjEcr.Detail[NumAxe - 1].Detail.Count <= 0 then Exit;

  // Init Sens et montants
  Sens   := GetSens( ObjEcr );
  MtEcrP := GetMontant( ObjEcr );
  MtEcrD := GetMontantDev( ObjEcr );

  // Détermination du taux
  if not (MtP = 0)
    then LeTaux := MtEcrP / MtP;
  if EnDevise(DEV.Code) then
  begin
    if not (MtD = 0)
      then LeTaux := MtEcrD / MtD;
  end;

  // Init totaux
  TotP := 0;
  TotD := 0;

  // Pour chaques lignes de ventilation, recalcul
  for NumV := 0 to (ObjEcr.Detail[NumAxe - 1].Detail.Count - 1) do
  begin
    ObjAna := ObjEcr.Detail[NumAxe - 1].Detail[NumV];

    // Calculs
    if MtP = 0 then
      begin
      LeTaux  := ObjAna.GetDouble('Y_POURCENTAGE') / 100.0 ;
      FDebit  := ObjEcr.GetValue('E_DEBIT');
      FCredit := ObjEcr.GetValue('E_CREDIT');
      end
    else
      begin
      FDebit := ObjAna.GetValue(Pf + '_DEBIT');
      FCredit := ObjAna.GetValue(Pf + '_CREDIT');
      end ;

    if ( Sens = 2 ) and ( FDebit <> 0 ) then
     begin // le sens de l'ecriture a changer, il faut inverser les montants de l'analytique
      FCredit := FDebit ;
      FDebit  := 0 ;
     end
      else // meme chose au debit
       if ( Sens = 1 ) and ( FCredit <> 0 ) then
        begin
         FDebit := FCredit ;
         FCredit  := 0 ;
        end ;
        
    Calc(FDebit, FCredit, TotP, LeTaux, Sens, V_PGI.OkDecV);
    ObjAna.PutValue(Pf + '_DEBIT', FDebit);
    ObjAna.PutValue(Pf + '_CREDIT', FCredit);
    FDebitDev := FDebit;
    FCreditDev := FCredit;

    {Devise}
    if EnDevise(DEV.Code) then
    begin
      if MtD = 0 then
        begin
        LeTaux  := ObjAna.GetDouble('Y_POURCENTAGE') / 100.0 ;
        FDebit  := ObjEcr.GetValue('E_DEBIT');
        FCredit := ObjEcr.GetValue('E_CREDIT');
        end
      else
        begin
        FDebitDev := ObjAna.GetValue(Pf + '_DEBITDEV');
        FCreditDev := ObjAna.GetValue(Pf + '_CREDITDEV');
        end ;
      Calc(FDebitDev, FCreditDev, TotD, LeTaux, Sens, DEV.Decimale);
    end;

    ObjAna.PutValue(Pf + '_DEBITDEV', FDebitDev);
    ObjAna.PutValue(Pf + '_CREDITDEV', FCreditDev);

    ObjAna.PutValue(Pf + '_TOTALECRITURE', MtEcrP);
    ObjAna.PutValue(Pf + '_TOTALDEVISE', MtEcrD);
    
  end;

  // Sur la dernière ligne
  if ObjEcr.Detail[NumAxe - 1].Detail.Count > 0 then
  begin
    NumV := ObjEcr.Detail[NumAxe - 1].Detail.Count - 1;
    ObjAna := ObjEcr.Detail[NumAxe - 1].Detail[NumV];

    if TotP <> MtEcrP then
    begin
      FDebit := ObjAna.GetValue(Pf + '_DEBIT');
      FCredit := ObjAna.GetValue(Pf + '_CREDIT');
      ReCalc(FDebit, FCredit, TotP, MtEcrp, Sens);
      ObjAna.PutValue(Pf + '_DEBIT', FDebit);
      ObjAna.PutValue(Pf + '_CREDIT', FCredit);
      if not EnDevise(DEV.Code) then
      begin
        FDebitDev := FDebit;
        FCreditDev := FCredit;
      end
      else
      begin
        FDebitDev := ObjAna.GetValue(Pf + '_DEBITDEV');
        FCreditDev := ObjAna.GetValue(Pf + '_CREDITDEV');
      end;
    end;
    if ((EnDevise(DEV.Code)) and (TotD <> MtEcrD)) then
    begin
      FDebitDev := ObjAna.GetValue(Pf + '_DEBITDEV');
      FCreditDev := ObjAna.GetValue(Pf + '_CREDITDEV');
      ReCalc(FDebitDev, FCreditDev, TotD, MtEcrD, Sens);
    end;
    ObjAna.PutValue(Pf + '_DEBITDEV', FDebitDev);
    ObjAna.PutValue(Pf + '_CREDITDEV', FCreditDev);
  end;
end;


procedure RecalculProrataAnalNEW(Pf: string; ObjEcr: TOB; NumA: integer; DEV: RDevise);
var
  MtP, MtD, MtE, TauxAn: Double;
  NumAxe: Integer;
begin
  for NumAxe := 1 to ObjEcr.Detail.Count do
    if ((NumA = 0) or (NumAxe = NumA)) then
      if ObjEcr.Detail[NumAxe - 1].Detail.Count > 0 then
      begin
        CalculTotalAnalNEW(Pf, ObjEcr, NumAxe, -1, MtP, MtD, MtE, TauxAn);
        Prorate1AnalNEW(Pf, ObjEcr, NumAxe, DEV, MtP, MtD, MtE, TauxAN);
        if GetSens(ObjEcr) <> GetSens(ObjEcr.Detail[NumAxe - 1].Detail[0])
          then InverseSensAnaNew(ObjEcr, NumAxe);
      end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 07/01/2004
Modifié le ... :   /  /
Description .. : Chargement en TOB de l'arbrescence anayltique depuis la
Suite ........ : base.
Suite ........ : Utilisé uniquement en cas de génération de l'analytique en
Suite ........ : mode "Report au détail"
Suite ........ : Cette arborescence est crée sur la tob des echéances
Suite ........ : d'origine passée en paramètres.
Mots clefs ... :
*****************************************************************}
procedure ChargeAnalytique( vTobPiece: TOB ; vDossier : String = '' );
var
  lInAxe: Integer;
  lInVentil: Integer;
  lInCpt: Integer;
  lTobAna: TOB;
  lTobAxe: TOB;
  lTobLigne: TOB;
  lStReq: string;
  lQAna: TQuery;
  lStWhere : string ;
  lTob  : Tob ;
begin

{$IFNDEF PGIIMMO}
  // vérification de la pièce (au moins une ligne)
  if vTobPiece.Detail.Count = 0 then
    Exit;

 lTOB := vTOBPiece.Detail[0] ;
 if (lTOB.GetString('E_MODESAISIE')='-') or (lTOB.GetString('E_MODESAISIE')='') then
   // Mode pièce
   lStWhere :=   ' Y_JOURNAL="'          + lTOB.GetString('E_JOURNAL')                     + '" ' +
             ' AND Y_EXERCICE="'         + lTOB.GetString('E_EXERCICE')                    + '" ' +
             ' AND Y_DATECOMPTABLE="'    + UsDateTime(lTOB.GetDateTime('E_DATECOMPTABLE')) + '" ' +
             ' AND Y_NUMEROPIECE='       + InttoStr(lTOB.GetInteger('E_NUMEROPIECE'))      +
             ' AND Y_QUALIFPIECE="'      + lTOB.GetString('E_QUALIFPIECE')                 + '" '
  else
   // Mode bordereau
   lStWhere :=   ' Y_JOURNAL="'          + lTOB.GetString('E_JOURNAL')                   + '" ' +
             ' AND Y_EXERCICE="'         + lTOB.GetString('E_EXERCICE')                  + '" ' +
             ' AND Y_DATECOMPTABLE>="'   + USDATETIME(DebutDeMois(lTOB.GetDateTime('E_DATECOMPTABLE'))) + '" ' +
             ' AND Y_DATECOMPTABLE<="'   + USDATETIME(FinDeMois(lTOB.GetDateTime('E_DATECOMPTABLE')))   + '" ' +
             ' AND Y_NUMEROPIECE='       + InttoStr(lTOB.GetInteger('E_NUMEROPIECE'))    +
             ' AND Y_QUALIFPIECE="'      + lTOB.GetString('E_QUALIFPIECE')               + '" ' ;


  // Chargement de toutes les lignes d'analytiques en une seule fois
  lTobAna := TOB.Create('VANAORIG', nil, -1);
  // SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
  lStReq  := 'SELECT ' + GetSelectAll('Y', True) + ', Y_BLOCNOTE FROM ANALYTIQ WHERE ' + lStWhere + ' ORDER BY Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL';

  if (Trim(vDossier)='') or (vDossier=V_PGI.SchemaName) then
    lTobAna.LoadDetailDBFromSql('ANALYTIQ', lStReq )
  else
    begin
    lQAna := OpenSelect( lStReq, vDossier ) ;
    if not lQAna.EOF then
      lTobAna.LoadDetailDB('ANALYTIQ', '', '', lQAna, False, True);
    Ferme(lQAna);
    end ;

  // Parcours de la pièce pour affectation des analytiques
  for lInCpt := 0 to vTobPiece.Detail.Count - 1 do
    begin
    lTobLigne := vTobPiece.Detail[lInCpt];

    // Allocation des axes
    if lTobLigne.Detail.count > 0 then
      lTobLigne.ClearDetail ;
    AlloueAxe( lTobLigne ) ;

    // Pour chaque axe...
    for lInAxe := 1 to MaxAxe do
      begin
      lTOBAxe := lTobLigne.Detail[ lInAxe - 1 ] ;

      // Parcours des lignes d'analytique
      for lInVentil := lTobAna.Detail.Count - 1 downto 0 do
        // Même numéro de ligne que ligne générale et même axe que axe est cours de traitement
        if ( lTobAna.Detail[ lInVentil ].GetValue('Y_NUMLIGNE') = lTobLigne.GetValue('E_NUMLIGNE') ) and
           ( lTobAna.Detail[ lInVentil ].GetValue('Y_AXE') = ('A' + IntToStr(lInAxe)) ) then
          lTobAna.Detail[ lInVentil ].ChangeParent( lTOBAxe, 0 ) ;
      end;

    end;

  // Libération mémoire
  lTobAna.ClearDetail; // normalement vide mais au cas ou...
  FreeAndNil(lTobAna);
  
{$ENDIF}

end;


{***********A.G.L.***********************************************
Auteur  ...... : SG6
Créé le ...... : 21/12/2004
Modifié le ... :   /  /
Description .. : Procedure qui permet de vérifier les quantités entre une
Suite ........ : ligne d'écriture et les eventuelles écritures analytiques
Suite ........ : associées
Suite ........ :
Suite ........ : FQ 14731
Mots clefs ... : CHECKAGE; QUANTITE, ANALYTIQUE, ECRITURE
*****************************************************************}
function CheckQuantite( ObjEcr: TOB ; vBoParle : Boolean = True ) : Boolean ;
var
  i, j: integer;
  qt1pourc, qt2pourc: double;
  qt1, qt2 : double;
  quantifiant1, quantifiant2: boolean;
  strMessage: string;
begin
  strMessage := '';
  //On parcourt tous les axes
  result := True ;
  if ObjEcr.Detail.Count = 0 then exit;
  for i := 0 to MaxAxe - 1 do
  begin
    //init des variables pour chaque axe
    if ObjEcr.Detail[i].Detail.Count > 0 then
    begin
      quantifiant1 := True;
      quantifiant2 := True;
      qt1pourc := 0;
      qt2pourc := 0;
      //Pour chaque ligne d"ecriture analytiques (par axe)
      for j := 0 to ObjEcr.Detail[i].Detail.Count - 1 do
      begin
        // FQ 15346 SBO Attention '' = '...' sur les qualifiant de pièces
        if (ObjEcr.GetValue('E_QUALIFQTE1') = '') or (ObjEcr.GetValue('E_QUALIFQTE1')='...')
          then quantifiant1 := quantifiant1 and ( (ObjEcr.Detail[i].Detail[j].GetValue('Y_QUALIFQTE1') = '') or (ObjEcr.Detail[i].Detail[j].GetValue('Y_QUALIFQTE1') = '...') )
          else quantifiant1 := quantifiant1 and (ObjEcr.Detail[i].Detail[j].GetValue('Y_QUALIFQTE1') = ObjEcr.GetValue('E_QUALIFQTE1'));
        if (ObjEcr.GetValue('E_QUALIFQTE2') = '') or (ObjEcr.GetValue('E_QUALIFQTE2') = '...')
          then quantifiant2 := quantifiant2 and ( (ObjEcr.Detail[i].Detail[j].GetValue('Y_QUALIFQTE2') = '') or (ObjEcr.Detail[i].Detail[j].GetValue('Y_QUALIFQTE2') = '...') )
          else quantifiant2 := quantifiant2 and (ObjEcr.Detail[i].Detail[j].GetValue('Y_QUALIFQTE2') = ObjEcr.GetValue('E_QUALIFQTE2'));
        // Fin FQ 15346
        qt1pourc := qt1pourc + ObjEcr.Detail[i].Detail[j].GetValue('Y_POURCENTQTE1');
        qt2pourc := qt2pourc + ObjEcr.Detail[i].Detail[j].GetValue('Y_POURCENTQTE2');
      end;
      //Génération du message
      qt1 := qt1pourc * ObjEcr.GetValue('E_QTE1') / 100;
      qt2 := qt2pourc * ObjEcr.GetValue('E_QTE2') / 100;
      if qt1 <> ObjEcr.GetValue('E_QTE1') then strMessage := strMessage + 'Les quantités qte1 saisies pour l''axe ' + IntToStr(i + 1) + ' sont incohérentes. (Ecriture : '
        + ' Qte1 = ' + strfmontant(ObjEcr.Getvalue('E_QTE1'), 15, 2, '', False) + ', Analytique : Qte1 = ' + strfmontant(qt1, 15, 2, '', False) + ' )' + Chr(13);
      if not (quantifiant1) then
        strMessage := strMessage + 'Le qualifiant de la quantité qt1 entre la ligne d''écriture et les lignes d''écritures analytiques de l''axe ' + IntTostr(i + 1) +
        ' doit être identique.' + Chr(13);
      if qt2 <> ObjEcr.GetValue('E_QTE2') then strMessage := strMessage + 'Les quantités qte2 saisies pour l''axe ' + IntToStr(i + 1) + ' sont incohérentes. (Ecriture : '
        + ' Qte2 = ' + strfmontant(ObjEcr.Getvalue('E_QTE2'), 15, 2, '', False) + ', Analytique : Qte2 = ' + strfmontant(qt2, 15, 2, '', False) + ' )' + Chr(13);
      if not (quantifiant2) then
        strMessage := strMessage + 'Le qualifiant de la quantité qt2 entre la ligne d''écriture et les lignes d''écritures analytiques de l''axe ' + IntTostr(i + 1) +
        ' doit être identique.' + Chr(13);
    end;
  end;

  // Mémorisation test // DEV3946
  if ObjEcr.GetNumChamp('CHECKQTE') < 0
    then ObjEcr.AddChampSupValeur('CHECKQTE', 'X' )
    else ObjEcr.PutValue('CHECKQTE', 'X' ) ;

  {$IFNDEF EAGLSERVER}
  if vBoParle and (strMessage <> '' ) then
    Hshowmessage('0;Incohérence des quantités saisies;' + strMessage + ';E;O;O;O;', '', '');
  {$ENDIF}

  result := strMessage = '' ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 16/02/2005
Modifié le ... :   /  /
Description .. : Fonction qui permet de renvoyer le premier_axe ventilé et le
Suite ........ : dernier_axe lorsqu'on utilise le mode analytique croisaxe
Mots clefs ... :
*****************************************************************}
function RecherchePremDerAxeVentil: TParamAxe;
var
  iCompteur: integer;
begin
  result.premier_axe := 0;
  result.dernier_axe := 0;

  if not GetAnaCroisaxe then Exit;

  // for iCompteur := 1 to 5 do BVE 10.05.07
  for iCompteur := 1 to MaxAxe do
  begin
    if GetParamSoc('SO_VENTILA' + IntToStr(iCompteur)) then
    begin
      if result.premier_axe = 0 then result.premier_axe := iCompteur;
      result.dernier_axe := iCompteur;
    end;
  end;
end;

////////////////////////////////////  Class Traitement Croisaxe ////////////////////////////////////

constructor TTraitementCroisaxe.Create(vA1, vA2, vA3, vA4, vA5, vImport: boolean; vTobEcr: TOB = nil);
begin
  cArrAxe[1] := vA1;
  cArrAxe[2] := vA2;
  cArrAxe[3] := vA3;
  cArrAxe[4] := vA4;
  cArrAxe[5] := vA5;

  cImport    := vImport ;

  cSoldeSection := TOB.Create('SOLDESECTION',nil,-1);
  cSoldeJournal := TOB.Create('SOLDEJOURNAUX',nil,-1);

  if vTobEcr = nil then exit;
  cTobEcr := TOB.Create(vTobEcr.nomTable, nil, -1);
  cTobNew := TOB.Create(vTobEcr.nomTable, nil, -1);
  InitClass(vTobEcr);

end;

constructor TTraitementCroisaxe.Create(vImport: boolean; vTobEcr: TOB = nil);
begin
  cArrAxe[1] := False;
  cArrAxe[2] := False;
  cArrAxe[3] := False;
  cArrAxe[4] := False;
  cArrAxe[5] := False;
  cImport    := vImport ;

  cSoldeSection := TOB.Create('SOLDESECTION',nil,-1);
  cSoldeJournal := TOB.Create('SOLDEJOURNAUX',nil,-1);
  
  if vTobEcr = nil then exit;
  cTobEcr := TOB.Create(vTobEcr.nomTable, nil, -1);
  cTobNew := TOB.Create(vTobEcr.nomTable, nil, -1);
  InitClass(vTobEcr);
end;

destructor TTraitementCroisaxe.Destroy;
begin
  FreeAndNil(cTobEcr);
end;

function TTraitementCroisaxe.PassageVentAnaClassiqueCroisaxe: TOB;
var
  n, p, k, i, j: Integer;
  S1, S2, S3, S4, S5: TOB;
  A1, A2, A3, A4, A5: TOB;
  T: TOB;
  MntDev, Pct, Mnt: Double;
  MntDevTmp, PctTmp, MntTmp: Double;
  lTobTmp  : TOB;
begin
  result := nil;
  cBooAnaCroisaxe := True;
  if cTobEcr.Detail.Count = 0 then Exit;

  lTobTmp := TOB.Create('$AXE', cTobNew, -1);
  lTobTmp.AddChampSupValeur('Y_AXE', 'A' + IntToStr(cIntPremAxe));


  A1 := GetTobAxe(0);
  A2 := GetTobAxe(1);
  A3 := GetTobAxe(2);
  A4 := GetTobAxe(3);
  A5 := GetTobAxe(4);

  Pct := 0;
  Mnt := 0;
  MntDev := 0;
  MntDevTmp := 0;
  MntTmp := 0;
  PctTmp := 0;
  for n := 0 to A1.Detail.Count - 1 do
  begin
    S1 := A1.Detail[n];
    if not assigned(A2) then
    begin
      PctTmp := NewAna(S1);
      T := cTobNew.Detail[0].Detail[cTobNew.Detail[0].Detail.Count - 1];
      MntTmp := CalculMontant(T, False);
      MntDevTmp := CalculMontant(T, True);
      Pct := Pct + PctTmp;
      Mnt := Mnt + MntTmp;
      MntDev := MntDev + MntDevTmp;
    end
    else
    begin
      for p := 0 to A2.Detail.Count - 1 do
      begin
        S2 := A2.Detail[p];
        if not assigned(A3) then
        begin
          PctTmp := NewAna(S1, S2);
          T := cTobNew.Detail[0].Detail[cTobNew.Detail[0].Detail.Count - 1];
          MntTmp := CalculMontant(T, False);
          MntDevTmp := CalculMontant(T, True);
          Pct := Pct + PctTmp;
          Mnt := Mnt + MntTmp;
          MntDev := MntDev + MntDevTmp;
        end
        else
          for k := 0 to A3.Detail.Count - 1 do
          begin
            S3 := A3.Detail[k];
            if not assigned(A4) then
            begin
              PctTmp := NewAna(S1, S2, S3);
              T := cTobNew.Detail[0].Detail[cTobNew.Detail[0].Detail.Count - 1];
              MntTmp := CalculMontant(T, False);
              MntDevTmp := CalculMontant(T, True);
              Pct := Pct + PctTmp;
              Mnt := Mnt + MntTmp;
              MntDev := MntDev + MntDevTmp;
            end
            else
              for i := 0 to A4.Detail.Count - 1 do
              begin
                S4 := A4.Detail[i];
                if not assigned(A5) then
                begin
                  PctTmp := NewAna(S1, S2, S3, S4);
                  T := cTobNew.Detail[0].Detail[cTobNew.Detail[0].Detail.Count - 1];
                  MntTmp := CalculMontant(T, False);
                  MntDevTmp := CalculMontant(T, True);
                  Pct := Pct + PctTmp;
                  Mnt := Mnt + MntTmp;
                  MntDev := MntDev + MntDevTmp;
                end
                else
                  for j := 0 to A5.Detail.Count - 1 do
                  begin
                    S5 := A5.Detail[j];
                    PctTmp := NewAna(S1, S2, S3, S4, S5);
                    T := cTobNew.Detail[0].Detail[cTobNew.Detail[0].Detail.Count - 1];
                    MntTmp := CalculMontant(T, False);
                    MntDevTmp := CalculMontant(T, True);
                    Pct := Pct + PctTmp;
                    Mnt := Mnt + MntTmp;
                    MntDev := MntDev + MntDevTmp;
                  end;
              end;
          end;
      end;
    end;
  end;
  VerifEquil(Mnt, MntTmp, MntDev, MntDevTmp, Pct, PctTmp);
  result := cTobNew;
end;



function TTraitementCroisaxe.GetPourCent(Section: TOB): Double;
begin
  if not Assigned(Section) then Result := 1
  else Result := Section.GetDouble('Y_POURCENTAGE') / 100.0;
end;

function TTraitementCroisaxe.CalculMontant(NewSec: TOB; vDevise: boolean): Double;
var
  lDblMontant: double;
  lIntDecimaleDevise: integer;
begin
  if vDevise then lDblMontant := NewSec.GetDouble('Y_TOTALDEVISE')
  else lDblMontant := NewSec.GetDouble('Y_TOTALECRITURE');

  if vDevise then lIntDecimaleDevise := cRecDevise.Decimale
  else lIntDecimaleDevise := V_PGI.OkDecV;

  Result := Arrondi(NewSec.GetDouble('Y_POURCENTAGE') * lDblMontant / 100.0, lIntDecimaleDevise);

  if cBooDebit then
  begin
    if vDevise then
    begin
      NewSec.SetDouble('Y_DEBITDEV', Result);
      NewSec.SetDouble('Y_CREDITDEV', 0);
    end
    else
    begin
      NewSec.SetDouble('Y_DEBIT', Result);
      NewSec.SetDouble('Y_CREDIT', 0);
      AjoutSoldeSection(NewSec);
    end;
  end
  else
  begin
    if vDevise then
    begin
      NewSec.SetDouble('Y_CREDITDEV', Result);
      NewSec.SetDouble('Y_DEBITDEV', 0);
    end
    else
    begin
      NewSec.SetDouble('Y_CREDIT', Result);
      NewSec.SetDouble('Y_DEBIT', 0);
      AjoutSoldeSection(NewSec);
    end;
  end;
end;

function TTraitementCroisaxe.NewAna(S1 : TOB; S2 : TOB = nil ; S3: TOB = nil; S4: TOB = nil; S5: TOB = nil): Double;
var
  T: TOB;
  lStrAxe: string;
  lIntAxe: integer;
begin
  {JP 04/11/05 : ADecimP au lieu de V_PGI.OkDecP pour être en accord avec la saisie analytique}
  Result := Arrondi(GetPourCent(S1) * GetPourCent(S2) * GetPourCent(S3) * GetPourCent(S4) * GetPourCent(S5) * 100.0, ADecimP);
  T := TOB.Create('ANALYTIQ', cTobNew.Detail[0], -1);
  T.InitValeurs ;
  CPutDefautAna(T);
  InitCommunObjAnalNEW(cTobEcr, T);
  T.SetDouble('Y_POURCENTAGE', Result);

  //Axe

  T.SetString('Y_AXE', 'A' + IntToStr(cIntPremAxe));

  //Sousplan

  lStrAxe := (S1.GetString('Y_AXE'))[2];
  T.SetString('Y_SOUSPLAN' + lStrAxe, S1.GetString('Y_SECTION'));
  if assigned(S2) then
  begin
    lStrAxe := (S2.GetString('Y_AXE'))[2];
    T.SetString('Y_SOUSPLAN' + lStrAxe, S2.GetString('Y_SECTION'));
  end;
  if assigned(S3) then
  begin
    lStrAxe := (S3.GetString('Y_AXE'))[2];
    T.SetString('Y_SOUSPLAN' + lStrAxe, S3.GetString('Y_SECTION'));
  end;
  if assigned(S4) then
  begin
    lStrAxe := (S4.GetString('Y_AXE'))[2];
    T.SetString('Y_SOUSPLAN' + lStrAxe, S4.GetString('Y_SECTION'));
  end;
  if assigned(S5) then
  begin
    lStrAxe := (S5.GetString('Y_AXE'))[2];
    T.SetString('Y_SOUSPLAN' + lStrAxe, S5.GetString('Y_SECTION'));
  end;

  for lIntAxe := 1 to MaxAxe do
  begin
    if cArrAxe[lIntAxe] and (T.GetString('Y_SOUSPLAN' + IntToStr(lIntAxe)) = '') then
      T.SetString('Y_SOUSPLAN' + IntToStr(lIntAxe), GetInfoCpta(AxeToFb('A' + IntToStr(lIntAxe))).Attente);
  end;

  T.SetString('Y_SECTION', T.GetString('Y_SOUSPLAN' + IntToStr(cIntPremAxe)));
  T.SetInteger('Y_NUMVENTIL', cIntNumVentil);

  Inc(cIntNumVentil);



end;

function TTraitementCroisaxe.GetTobAxe(Ind: Integer): TOB;
var
  n: Byte;
begin
  result := nil;
  for n := Ind to cTobEcr.Detail.Count - 1 do
  begin
    if cTobEcr.Detail[n].Detail.Count <> 0 then
    begin
      Result := cTobEcr.Detail[n];
      break;
    end;
  end;
end;

procedure TTraitementCroisaxe.VerifEquil(Mnt, MntTmp, MntDev, MntDevTmp, Pct, PctTmp: Double);
var
  T: TOB;
  lDblTotalEcr: double;
  lDblTotalEcrDev: double;
begin
  T := cTobNew.Detail[0].Detail[cTobNew.Detail[0].Detail.Count - 1];
  //Equilibrage pourcentage
  if Pct <> 100 then T.SetDouble('y_Pourcentage', 100 - (Pct - PctTmp));

  lDblTotalEcr := cTobEcr.GetDouble('E_CREDIT') + cTobEcr.GetDouble('E_DEBIT');
  lDblTotalEcrDev := cTobEcr.GetDouble('E_CREDITDEV') + cTobEcr.GetDouble('E_DEBITDEV');
  //Equilibrage en devise de tenue
  if Mnt <> lDblTotalEcr then
    if cBooDebit then
      T.SetDouble('Y_DEBIT', lDblTotalEcr - (Mnt - MntTmp))
    else
      T.SetDouble('Y_CREDIT', lDblTotalEcr - (Mnt - MntTmp));



  //Equilibrage montant en devise
  if MntDev <> lDblTotalEcrDev then
    if cBooDebit then
      T.SetDouble('Y_DEBIT', lDblTotalEcrDev - (MntDev - MntDevTmp))
    else
      T.SetDouble('Y_CREDIT', lDblTotalEcrDev - (MntDev - MntDevTmp));
end;

procedure TTraitementCroisaxe.InitClass(vTobEcr: TOB);
var
  lIntCpteur: integer;
  lIntAxe: integer;
begin
  // Gestion TOB
  cTobEcr.Dupliquer(vTobEcr, True, True);

  if cTobEcr.nomTable = 'ECRITURE' then
  begin
    cTobNew.Dupliquer(vTobEcr, True, True);
    cTobNew.ClearDetail;
  end
  else if cTobEcr.nomTable = 'ANALYTIQ' then
  begin
    cTobNew.Dupliquer(vTobEcr, True, True);
  end
  else
  begin
    TOB.Create('PIECE',cTobNew, -1);
    cTobNew.Detail[0].Dupliquer(vTobEcr, True, True);
  end;

  cIntPremAxe := 0;

  for lIntCpteur := 1 to MaxAxe do
  begin
    if cArrAxe[lIntCpteur] then {JP 02/08/07 : au lieu de cArrAxe[1] !!!}
      if cIntPremAxe = 0 then cIntPremAxe := lIntCpteur;
  end;

  //On efface les axes n'etant pas paramétré en mode croisaxe
  if (cTobEcr.nomTable = 'ECRITURE') and not GetAnaCroisaxe then
  begin
    for lIntCpteur := cTobEcr.Detail.Count - 1 downto 0 do
    begin
      if cTobEcr.Detail[lIntCpteur].Detail.Count <> 0 then
      begin
        lIntAxe := ValeurI((cTobEcr.Detail[lIntCpteur].Detail[0].GetString('Y_AXE'))[2]);
        if not cArrAxe[lIntAxe] then
        begin
          //Maj si on est pas en import (ComSx) on met a jour les soldes
          if not cImport then MajSoldeSectionDelete(lIntCpteur);
          cTobEcr.Detail[lIntCpteur].Free;
        end;
      end;
    end;
  end;


  //Gestion devise ecriture
  if cTobEcr.nomTable = 'ECRITURE' then
    begin
    cRecDevise.Code := cTobEcr.GetString('E_DEVISE') ;
    cBooDebit := cTobEcr.GetDouble('E_DEBIT') <> 0.0 ;
    end
  else if cTobEcr.nomTable = 'ANALYTIQ' then
    begin
    cRecDevise.Code := cTobEcr.GetString('Y_DEVISE');
    cBooDebit := cTobEcr.GetDouble('Y_DEBIT') <> 0.0;
    end ;
  GetInfosDevise(cRecDevise);

  cIntNumVentil := 0;

end;

procedure TTraitementCroisaxe.SetTobEcr(vTobEcr: TOB);
begin
  if vTobEcr = nil then exit;
  if Assigned(cTobEcr) then FreeAndNil(cTobEcr);
  if Assigned(cTobNew) then FreeAndNil(cTobNew);

  cTobEcr := TOB.Create(vTobEcr.nomTable, nil, -1);
  cTobNew := TOB.Create(vTobEcr.nomTable, nil, -1);
  InitClass(vTobEcr);
end;

function TTraitementCroisaxe.PassageAnaPurClassiqueCroisaxe: TOB;
var
  lIntCpteur: integer;
  lIntAxe : integer;
begin
  cBooAnaCroisaxe := True;
  result := nil;
  if cTobEcr = nil then exit;

  lIntAxe := ValeurI(Copy(cTobNew.GetString('Y_AXE'), 2, 1));
  cTobNew.SetString('Y_SOUSPLAN' + IntToStr(lIntAxe), cTobNew.GetString('Y_SECTION'));
  cTobNew.SetString('Y_AXE', '');
  cTobNew.SetString('Y_SECTION', '');

  for lIntCpteur := 1 to MaxAxe do
  begin
    if GetParamSoc('SO_VENTILA' + IntToStr(lIntCpteur)) then
    begin
      if cTobNew.GetString('Y_SOUSPLAN' + IntToStr(lIntCpteur)) = '' then
        cTobNew.SetString('Y_SOUSPLAN' + IntToStr(lIntCpteur), GetInfoCpta(AxeToFb('A' + IntToStr(lIntCpteur))).Attente);

      if cTobNew.GetString('Y_AXE') = '' then
      begin
        cTobNew.SetString('Y_AXE', 'A' + IntToStr(lIntCpteur));
        cTobNew.SetString('Y_SECTION', cTobNew.GetString('Y_SOUSPLAN' + IntToStr(lIntCpteur)));
      end;
    end;
  end;
  result := cTobNew;

  //Maj Solde
  AjoutSoldeSection(cTobNew);
end;

procedure TTraitementCroisaxe.DeleteInsertAna;
begin

  //delete
  if cTobEcr.NomTable = 'ECRITURE' then
  begin
    ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_JOURNAL="'+ cTobEcr.GetString('E_JOURNAL') + '" AND Y_EXERCICE="' + ctobEcr.GetString('E_EXERCICE') + '" AND Y_DATECOMPTABLE="' + usdatetime(ctobEcr.GetDateTime('E_DATECOMPTABLE')) +
                   '" AND Y_NUMEROPIECE=' + IntToStr(ctobEcr.GetInteger('E_NUMEROPIECE')) +
                   ' AND Y_NUMLIGNE=' + IntToStr(ctobEcr.GetInteger('E_NUMLIGNE')) +
                   ' AND Y_QUALIFPIECE="' + ctobEcr.GetString('E_QUALIFPIECE') + '"');
  end
  else if  cTobEcr.NomTable = 'ANALYTIQ' then
  begin
     cTobEcr.DeleteDB;
  end
  else
  begin
    ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_JOURNAL="'+ cTobEcr.Detail[0].GetString('Y_JOURNAL') + '" AND Y_EXERCICE="' + cTobEcr.Detail[0].GetString('Y_EXERCICE') + '" AND Y_DATECOMPTABLE="' + usdatetime(cTobEcr.Detail[0].GetDateTime('Y_DATECOMPTABLE')) +
                   '" AND Y_NUMEROPIECE=' + IntToStr(cTobEcr.Detail[0].GetInteger('Y_NUMEROPIECE')) +
                   ' AND Y_AXE="' + cTobEcr.Detail[0].GetString('Y_AXE') +
                   '" AND Y_NUMLIGNE=' + IntToStr(cTobEcr.Detail[0].GetInteger('Y_NUMLIGNE')) +
                   ' AND Y_QUALIFPIECE="' + cTobEcr.Detail[0].GetString('Y_QUALIFPIECE') + '"');
  end;

  //Insert
  if cTobEcr.nomTable = 'ECRITURE' then
  begin
    cTobNew.InsertDBByNivel(False,2,2);
  end
  else if cTobEcr.nomTable = 'ANALYTIQ' then
  begin
    cTobNew.InsertDB(nil);
  end
  else
  begin
    cTobNew.InsertDBByNivel(False,2,2);
//    MajJournalAnaTob(cTobNew.Detail[lIntPiece].Detail[lIntAna], False, True);
  end;
end;

procedure TTraitementCroisaxe.MajSoldeSectionDelete(vAxe : integer);
var
  lIntCpteur  : integer;
begin
  if assigned(cTobEcr.Detail[vAxe]) then
  begin
    for lIntCpteur := 0 to cTobEcr.Detail[vAxe].Detail.Count - 1 do MajSoldeSectionTOB(cTobEcr.Detail[vAxe].Detail[lIntCpteur], False);
  end;
end;


function TTraitementCroisaxe.PassageVentCroisaxeAnaClassique: TOB;
var
  lIntCpteur : integer;
  lIntSousPlan : integer;
  lStrSousPlan  : string;
  lTobAxe    : TOB;
  lTobNewAxe    : TOB;
  lTobAna    : TOB;
  lTobNewAna  : TOB;
  lArrayNumVentil : array[1..5] of longint;
begin
  cBooAnaCroisaxe := False;
  result := nil;
  if cTobEcr.Detail.Count = 0 then Exit;

  FillChar(lArrayNumVentil,SizeOf(lArrayNumVentil), #0);

  AlloueAxe(cTobNew);

  lTobAxe := GetTobAxe(0);
  if not assigned(lTobAxe) then exit;

  for lIntCpteur := 0 to lTobAxe.Detail.Count - 1 do
  begin
    //Pour tous les sousplans
    lTobAna := lTobAxe.Detail[lIntCpteur];
    for lIntSousPlan := 1 to MaxAxe do
    begin
      lTobNewAxe := cTobNew.Detail[lIntSousPlan-1];

      lStrSousPlan := lTobAna.GetString('Y_SOUSPLAN' + IntToStr(lIntSousPlan));
      //Remplissage de cArrAxe
      if (not cImport) and (not lIntCpteur = 0) then
      begin
        if lStrSousPlan <> '' then cArrAxe[lIntsousPlan] := True;
      end;
(*  ajout me*)
      // ajout me  Cas Import Comsx du fichier croisaxe vers dossier non croisaxe
      if (lStrSousPlan <> '') and  cArrAxe[lIntsousPlan] then
      begin
        lTobNewAna := lTobNewAxe.FindFirst(['Y_SECTION'],[lStrSousPlan],True);
        if lTobNewAna <> nil then
        begin
          //On met à jour la ligne d'ecriture ana
          lTobNewAna.SetDouble('Y_POURCENTAGE', lTobNewAna.GetDouble('Y_POURCENTAGE') + lTobAna.GetDouble('Y_POURCENTAGE'));
          lTobNewAna.SetDouble('Y_CREDIT', lTobNewAna.GetDouble('Y_CREDIT') + lTobAna.GetDouble('Y_CREDIT'));
          lTobNewAna.SetDouble('Y_DEBIT', lTobNewAna.GetDouble('Y_DEBIT') + lTobAna.GetDouble('Y_DEBIT'));
          lTobNewAna.SetDouble('Y_CREDITDEV', lTobNewAna.GetDouble('Y_CREDITDEV') + lTobAna.GetDouble('Y_CREDITDEV'));
          lTobNewAna.SetDouble('Y_DEBITDEV', lTobNewAna.GetDouble('Y_DEBITDEV') + lTobAna.GetDouble('Y_DEBITDEV'));
        end
        else
        begin
          //On cree la ligne d'analytique
          lTobNewAna := TOB.Create('ANALYTIQ', lTobNewAxe, -1);
          lTobNewAna.InitValeurs ;
          CPutDefautAna(lTobNewAna);
          InitCommunObjAnalNEW(cTobEcr, lTobNewAna);

          //Axe
          lTobNewAna.SetString('Y_AXE', 'A' + IntToStr(lIntSousPlan));

          //Section
          lTobNewAna.SetString('Y_SECTION', lStrSousPlan);

          //Pourcentage
          lTobNewAna.SetDouble('Y_POURCENTAGE', lTobAna.GetDouble('Y_POURCENTAGE'));

          //Debit
          lTobNewAna.SetDouble('Y_DEBIT', lTobAna.GetDouble('Y_DEBIT'));
          lTobNewAna.SetDouble('Y_DEBITDEV', lTobAna.GetDouble('Y_DEBITDEV') );

          //Credit
          lTobNewAna.SetDouble('Y_CREDIT', lTobAna.GetDouble('Y_CREDIT'));
          lTobNewAna.SetDouble('Y_CREDITDEV', lTobAna.GetDouble('Y_CREDITDEV'));

          //NumVentil
          Inc(lArrayNumVentil[lIntSousPlan]);
          lTobNewAna.SetInteger('Y_NUMVENTIL', lArrayNumVentil[lIntSousPlan]);
        end;
      end;
    end;
  end;

  VerifEquilCroisaxe;
  result := cTobNew;

end;

procedure TTraitementCroisaxe.VerifEquilCroisaxe;
var
  lIntAxe : integer;
  lIntAna : integer;
  lDblMontantCumul    : double;
  lDblPourcentCumul   : double;
  lDblMontantDevCumul : double;
  vTobAna  : TOB;
begin

  for lIntAxe := 1 to MaxAxe do
  begin
    lDblMontantCumul    := 0;
    lDblPourcentCumul   := 0;
    lDblMontantDevCumul := 0;

    vTobAna := nil;

    if cTobNew.Detail[lIntAxe-1].Detail.Count = 0 then Continue;

    for lIntAna := 0 to cTobNew.Detail[lIntAxe-1].Detail.Count - 1 do
    begin
      vTobAna := cTobNew.Detail[lIntAxe-1].Detail[lIntAna];
      lDblMontantCumul := vTobAna.GetDouble('Y_CREDIT') + vTobAna.GetDouble('Y_DEBIT') + lDblMontantCumul;
      lDblMontantDevCumul := vTobAna.GetDouble('Y_CREDITDEV') + vTobAna.GetDouble('Y_DEBITDEV') + lDblMontantDevCumul;
      lDblPourcentCumul := vTobAna.GetDouble('Y_POURCENTAGE') + lDblPourcentCumul;
      if lIntAna <> (cTobNew.Detail[lIntAxe-1].Detail.Count - 1) then AjoutSoldeSection(vTobAna);
    end;

    //Equilibrage montant
    if (cTobNew.GetDouble('E_CREDIT') + cTobNew.GetDouble('E_DEBIT')) <> lDblMontantCumul then
      if cBooDebit then
        vTobAna.SetDouble('Y_DEBIT', cTobNew.GetDouble('E_TOTALECRITURE') - (lDblMontantCumul - vTobAna.GetDouble('Y_DEBIT')))
      else
        vTobAna.SetDouble('Y_CREDIT', cTobNew.GetDouble('E_TOTALECRITURE') - (lDblMontantCumul - vTobAna.GetDouble('Y_CREDIT')));

    //Equilibrage devise
    if (cTobNew.GetDouble('E_CREDITDEV') + cTobNew.GetDouble('E_DEBITDEV')) <> lDblMontantDevCumul then
      if cBooDebit then
        vTobAna.SetDouble('Y_DEBIT', cTobNew.GetDouble('E_TOTALDEVISE') - (lDblMontantDevCumul - vTobAna.GetDouble('Y_DEBIT')))
      else
        vTobAna.SetDouble('Y_CREDIT', cTobNew.GetDouble('E_TOTALDEVISE') - (lDblMontantDevCumul - vTobAna.GetDouble('Y_CREDIT')));

    //Equilibrage pourcent
    if lDblPourcentCumul <> 100.00 then
        vTobAna.SetDouble('Y_POURCENTAGE', 100.00 - (lDblPourcentCumul - vTobAna.GetDouble('Y_POURCENTAGE')));

    AjoutSoldeSection(vTobAna);
  end;
end;

function TTraitementCroisaxe.PassageAnaPurCroisaxeClassique: TOB;
var
  lIntAxe : integer;
  lStrSousPlan  : string;
  lBooDupliquer : boolean;
  lIntAna  : integer;
  lIntPiece : integer;
  lTobAna  : TOB;
  lIntCpteur  : integer;
  lIntNumPiece : integer;
  lStrJal : string;
  lDblDeb : double;
  lDblCre : double;
begin
  lBooDupliquer := False;

  for lIntAxe := 1 to MaxAxe do
  begin
    lStrSousPlan := cTobNew.Detail[0].Detail[0].GetString('Y_SOUSPLAN' + IntToStr(lIntAxe));
    //Remplissage de cArrAxe
    if lStrSousPlan <> '' then
    begin
      cArrAxe[lIntAxe] := True;
      if lBooDupliquer then
      begin
        TOB.Create('PIECE' + IntToStr(lIntAxe + 1),cTobNew, -1);
        cTobNew.Detail[cTobNew.Detail.Count - 1].Dupliquer(cTobNew.Detail[0],True,True);
      end;
      lBooDupliquer := True;
    end;
  end;

  //Maj Champs
  lIntPiece := 0;
  for lIntAxe := 1 to MaxAxe do
  begin
    if cArrAxe[lIntAxe] then
    begin
      lTobAna := nil;
      lStrJal := RechercheJournalOD(lIntAxe);
      lIntNumPiece := GetNewNumJal(lStrJal,True,cTobNew.Detail[0].GetDateTime('Y_DATECOMPTABLE'));
      lDblDeb := 0.0;
      lDblCre := 0.0;

      for lIntAna := 0 to cTobNew.Detail[lIntPiece].Detail.Count - 1 do
      begin
        lTobAna := cTobNew.Detail[lIntPiece].Detail[lIntAna];
        lTobAna.SetString('Y_AXE', 'A' + IntToStr(lIntAxe));
        lTobAna.SetString('Y_SECTION', cTobNew.Detail[0].Detail[lIntAna].GetString('Y_SOUSPLAN' + IntToStr(lIntAxe)));
        lTobAna.SetString('Y_JOURNAL', lStrJal);
        lTobAna.SetInteger('Y_NUMEROPIECE', lIntNumPiece);
        lDblCre := lDblCre + lTobAna.GetDouble('Y_CREDIT');
        lDblDeb := lDblDeb + lTobAna.GetDouble('Y_DEBIT');
        AjoutSoldeSection(lTobAna);
      end;
      if lTobAna <> nil then   AjoutSoldeJournal(lTobAna, lDblDeb, lDblCre);
    end;

    for lIntCpteur := 0 to cTobNew.Detail[0].Detail.Count - 1 do
    begin
      cTobNew.Detail[0].Detail[lIntCpteur].SetString('Y_SOUSPLAN' + IntToStr(lIntAxe), '');
    end;

    Inc(lIntPiece);
  end;

  result := cTobNew;
end;

{JP 07/11/05 : On croise aussi le paramétrage, à tout le moins les Ventilations Types
{---------------------------------------------------------------------------------------}
procedure TTraitementCroisaxe.CroiseVentilTypeEtAutres;
{---------------------------------------------------------------------------------------}
var
  Tob1  : TOB;
  Tob2  : TOB;
  Tob3  : TOB;
  Tob4  : TOB;
  Tob5  : TOB;
  TobN  : TOB;
  TobC  : TOB;
  n     : Integer;
  Where : string;
  Cpte  : string;
  Nat   : string;
begin
  Tob1 := nil;
  Tob2 := nil;
  Tob3 := nil;
  Tob4 := nil;
  Tob5 := nil;

  {Création des tobs contenant les ventilations par Axes, si les axes sont traités}
  if cArrAxe[1] then begin
    Tob1 := TOB.Create('$VENTIL', nil, -1);
    Where := 'V_NATURE LIKE "__1" ';
  end;
  if cArrAxe[2] then begin
    Tob2 := TOB.Create('$VENTIL', nil, -1);
    if Where = '' then Where := 'V_NATURE LIKE "__2" '
                  else Where := Where + 'OR V_NATURE LIKE "__2" ';
  end;
  if cArrAxe[3] then begin
    Tob3 := TOB.Create('$VENTIL', nil, -1);
    if Where = '' then Where := 'V_NATURE LIKE "__3" '
                  else Where := Where + 'OR V_NATURE LIKE "__3" ';
  end;
  if cArrAxe[4] then begin
    Tob4 := TOB.Create('$VENTIL', nil, -1);
    if Where = '' then Where := 'V_NATURE LIKE "__4" '
                  else Where := Where + 'OR V_NATURE LIKE "__4" ';
  end;
  if cArrAxe[5] then begin
    Tob5 := TOB.Create('$VENTIL', nil, -1);
    if Where = '' then Where := 'V_NATURE LIKE "__5" '
                  else Where := Where + 'OR V_NATURE LIKE "__5" ';
  end;

  {Tob contenant les nouvelles ventilations}
  TobN := TOB.Create('$VENTIL', nil, -1);
  {Tob contenant les comptes, natures et sociétés}
  TobC := TOB.Create('$VENTIL', nil, -1);
  try
    {Chargement des sections d'attente}
    SetSectAtt;
    
    {Chargement des différents comptes, natures et sociétés dans la table}
    TobC.LoadDetailFromSQL('SELECT DISTINCT(SUBSTRING(V_NATURE, 1, 2)||"-@-"||V_COMPTE) AS CHP FROM VENTIL WHERE ' + Where);

    {Chargement des ventilations par axes}
    if Assigned(Tob1) then Tob1.LoadDetailFromSQL('SELECT * FROM VENTIL WHERE V_NATURE LIKE "__1" ORDER BY V_NATURE, V_COMPTE');
    if Assigned(Tob2) then Tob2.LoadDetailFromSQL('SELECT * FROM VENTIL WHERE V_NATURE LIKE "__2" ORDER BY V_NATURE, V_COMPTE');
    if Assigned(Tob3) then Tob3.LoadDetailFromSQL('SELECT * FROM VENTIL WHERE V_NATURE LIKE "__3" ORDER BY V_NATURE, V_COMPTE');
    if Assigned(Tob4) then Tob4.LoadDetailFromSQL('SELECT * FROM VENTIL WHERE V_NATURE LIKE "__4" ORDER BY V_NATURE, V_COMPTE');
    if Assigned(Tob5) then Tob5.LoadDetailFromSQL('SELECT * FROM VENTIL WHERE V_NATURE LIKE "__5" ORDER BY V_NATURE, V_COMPTE');

    for n := 0 to TobC.Detail.Count - 1 do begin
      {Récupération de la nature du compte et de la société}
      Nat  := Copy(TobC.Detail[n].GetString('CHP'), 1, 2);
      Cpte := Copy(TobC.Detail[n].GetString('CHP'), 6, 35);

      {Calcul d'une nouvelle ventilation pour la société, la nature et le compte courants}
      GetLesPourcentage(Cpte, Nat, TobN, Tob1, Tob2, Tob3, Tob4, Tob5);
    end;

    {Si la suppression des anciennes ventilations s'est bien passée ...}
    if SupprimeVentiType(False) then
      {... On insère les nouvelles dans la table}
      TobN.InsertDb(nil);

  finally
    if Assigned(Tob1) then FreeAndNil(Tob1);
    if Assigned(Tob2) then FreeAndNil(Tob2);
    if Assigned(Tob3) then FreeAndNil(Tob3);
    if Assigned(Tob4) then FreeAndNil(Tob4);
    if Assigned(Tob5) then FreeAndNil(Tob5);
    if Assigned(TobC) then FreeAndNil(TobC);
    if Assigned(TobN) then FreeAndNil(TobN);
  end;
end;

{JP 07/11/05 : Calcul de la nouvelle ventilation pour une nature et un compte donnés
{---------------------------------------------------------------------------------------}
procedure TTraitementCroisaxe.GetLesPourcentage(Cpte, Nat : string; var TobN, Tob1, Tob2, Tob3, Tob4, Tob5 : TOB);
{---------------------------------------------------------------------------------------}
var
  T1 : TOB;
  T2 : TOB;
  T3 : TOB;
  T4 : TOB;
  T5 : TOB;
  D1 : TOB;
  D2 : TOB;
  D3 : TOB;
  D4 : TOB;
  D5 : TOB;
  Sec : string;
  SP1 : string;
  SP2 : string;
  SP3 : string;
  SP4 : string;
  SP5 : string;
  Qt1 : Double;
  Qt2 : Double;
  Pct : Double;
  CV  : array[1..MaxAxe] of Byte;
  Num : Integer;

    {Détermination du rang des axes paramétrés :
     si Axe 1 et 2 paramétrés Rang 1 = Tob1 et Rang 2 Tob2
     si Axe 2 et 4 paramétrés Rang 1 = Tob2 et Rang 2 Tob4
    {-------------------------------------------------------------------------}
    function RetourneTobToInd(Axe : Byte) : TOB;
    {-------------------------------------------------------------------------}
    begin
      case CV[Axe] of
        1 : Result := Tob1;
        2 : Result := Tob2;
        3 : Result := Tob3;
        4 : Result := Tob4;
        5 : Result := Tob5;
        else
          Result := nil;
      end;
    end;

    {Récupération de la section d'attente en fonction du rang de l'axe traité
    {-------------------------------------------------------------------------}
    procedure SectionAttSurAxe(Ind : Byte);
    {-------------------------------------------------------------------------}
    var
      k : Byte;
    begin
      {Récupération du numéro d'axe en fonction du rang}
      k := CV[Ind];
      if k = 0 then Exit;

      case k of
        1 : begin
             SP1 := SectAtt1;
             {Si on passe par là lorsque Ind est à 1 (c'est à dire que l'on est sur le premier axe
              paramétré) alors le champs v_section contiendra la section d'attente de cette axe}
             if Ind = 1 then Sec := SectAtt1;
           end;
        2 : begin
             SP2 := SectAtt2;
             {Si on passe par là lorsque Ind est à 1 (c'est à dire que l'on est sur le premier axe
              paramétré) alors le champs v_section contiendra la section d'attente de cette axe}
             if Ind = 1 then Sec := SectAtt2;
           end;
        3 : begin
             SP3 := SectAtt3;
             {Si on passe par là lorsque Ind est à 1 (c'est à dire que l'on est sur le premier axe
              paramétré) alors le champs v_section contiendra la section d'attente de cette axe}
             if Ind = 1 then Sec := SectAtt3;
           end;
        4 : begin
             SP4 := SectAtt4;
             {Si on passe par là lorsque Ind est à 1 (c'est à dire que l'on est sur le premier axe
              paramétré) alors le champs v_section contiendra la section d'attente de cette axe}
             if Ind = 1 then Sec := SectAtt4;
           end;
        5 : begin
             SP5 := SectAtt5;
             {Si on passe par là lorsque Ind est à 1 (c'est à dire que l'on est sur le premier axe
              paramétré) alors le champs v_section contiendra la section d'attente de cette axe}
             if Ind = 1 then Sec := SectAtt5;
           end;
      end;
    end;

    {-------------------------------------------------------------------------}
    procedure InitVariables(Ind : Byte);
    {-------------------------------------------------------------------------}
    begin
      {La section n'est réinitialisée qu'au début du traitement ou si l'on se
       déplace sur le premier axe de paramétré}
      if (Ind <= 1) then Sec := '';
      {On réinitialise les sous-plans que si on se déplace sur un axe "inférieur" au sous-plan}
      if (CV[Ind] <= 1) then SP1 := '';
      if (CV[Ind] <= 2) then SP2 := '';
      if (CV[Ind] <= 3) then SP3 := '';
      if (CV[Ind] <= 4) then SP4 := '';
      SP5 := '';
      Qt1 := 0;
      Qt2 := 0;
      Pct := 0;
    end;

    {-------------------------------------------------------------------------}
    procedure NouvelleVentilation(Ind : Byte);
    {-------------------------------------------------------------------------}
    var
      TobD : TOB;
    begin
      TobD := TOB.Create('VENTIL', TobN, -1);
      TobD.SetString('V_NATURE'   , Nat + IntToStr(CV[1]));
      TobD.SetString('V_COMPTE'   , Cpte);
      TobD.SetString('V_SECTION'  , Sec);
      TobD.SetString('V_SOCIETE'  , V_PGI.CodeSociete);
      TobD.SetString('V_SOUSPLAN1', SP1);
      TobD.SetString('V_SOUSPLAN2', SP2);
      TobD.SetString('V_SOUSPLAN3', SP3);
      TobD.SetString('V_SOUSPLAN4', SP4);
      TobD.SetString('V_SOUSPLAN5', SP5);
      TobD.SetDouble('V_MONTANT'  , 0);
      TobD.SetDouble('V_TAUXQT1'  , Arrondi(QT1 * 100, ADecimP));
      TobD.SetDouble('V_TAUXQT2'  , Arrondi(QT2 * 100, ADecimP));
      TobD.SetDouble('V_TAUXMONTANT'  , Arrondi(Pct * 100, ADecimP));
      TobD.SetInteger('V_NUMEROVENTIL', Num);
      {Réinitialisation des variables en fonction du niveau de récursivité en cours}
      InitVariables(Ind);
      {Incrémentation du numéro de ventilation}
      Inc(Num);
    end;

    {Calcul des valeurs de la nouvelle ventilation
    {-------------------------------------------------------------------------}
    procedure MajVariables(Ind : Byte; DD : TOB);
    {-------------------------------------------------------------------------}
    begin
      {Récupération de la section que l'on place dans le sous-plan correspondant
       au numéro de l'axe traité}
      case CV[Ind] of
        1 : SP1 := DD.GetString('V_SECTION');
        2 : SP2 := DD.GetString('V_SECTION');
        3 : SP3 := DD.GetString('V_SECTION');
        4 : SP4 := DD.GetString('V_SECTION');
        5 : SP5 := DD.GetString('V_SECTION');
        else
          Exit;
      end;

      {Si on est au rang 1 de la récursivité, le champ V_SECTION de la nouvelle
       ventilation aura alors la valeur actuelle de V_SECTION}
      if Ind = 1 then Sec := DD.GetString('V_SECTION');

      {Calcul des quantité et du Pourcentage}
      if QT1 > 0 then QT1 := QT1 * DD.GetDouble('V_TAUXQTE1') / 100
                 else QT1 := DD.GetDouble('V_TAUXQTE1') / 100;
      if QT2 > 0 then QT2 := QT2 * DD.GetDouble('V_TAUXQTE2') / 100
                 else QT2 := DD.GetDouble('V_TAUXQTE2') / 100;
      if Pct > 0 then Pct := Pct * DD.GetDouble('V_TAUXMONTANT') / 100
                 else Pct := DD.GetDouble('V_TAUXMONTANT') / 100;
    end;

    {-------------------------------------------------------------------------}
    procedure BoucleSur(Ind : Byte);
    {-------------------------------------------------------------------------}

         {Récupération de la Tob (Mère ou fille) en fonction du rang de la récursivité
         {-----------------------------------------------------------}
         function RetourneTob(MaitOk : Boolean = False) : TOB;
         {-----------------------------------------------------------}
         begin
           Result := nil;
           case Ind of
             1 : if MaitOk then Result := T1
                           else Result := T1.FindNext(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[Ind]), Cpte], True);
             2 : if MaitOk then Result := T2
                           else Result := T2.FindNext(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[Ind]), Cpte], True);
             3 : if MaitOk then Result := T3
                           else Result := T3.FindNext(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[Ind]), Cpte], True);
             4 : if MaitOk then Result := T4
                           else Result := T4.FindNext(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[Ind]), Cpte], True);
             5 : if MaitOk then Result := T5
                           else Result := T5.FindNext(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[Ind]), Cpte], True);
           end;
         end;
    var
      TT : TOB;
      DD : TOB;
      Ol1 : Double;
      Ol2 : Double;
      OlP : Double;
    begin
      DD := nil;
      {Fin de la récursivité}
      if CV[Ind] > MaxAxe then begin
        {Création d'une nouvelle ventilation}
        NouvelleVentilation(Ind - 1);
        {On sort de la récursivité}
        Exit;
      end;

      {Récupération de la tob contenant les ventilations d'un axe en fonction du rang de la récursivité}
      TT := RetourneTob(True);
      {Récupération de la première ventilation}
      if Assigned(TT) then
        DD := TT.FindFirst(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[Ind]), Cpte], True);

      {S'il n'y a pas de tob fille et que l'axe doit être traité, c'est que la ventilation n'a pas été
       paramétrée sur cet axe. On va donc mettre la section d'attente dans le champ Y_SOUSPLAN qui
       correspond au numéro de l'axe}
      if not Assigned(DD) and cArrAxe[CV[Ind]] then begin
        {Récupération de la section d'attente pour l'axe en cours}
        SectionAttSurAxe(Ind);
        {Nouveau niveau de récursivité}
        BoucleSur(Ind + 1);
      end

      {L'axe n'est pas paramétré}
      else if not cArrAxe[CV[Ind]] then begin
        {Création d'une nouvelle ventilation}
        NouvelleVentilation(Ind - 1);
      end

      {Il y a une ventilation type sur l'axe en cours}
      else begin
        Ol1 := 0;
        Ol2 := 0;
        OlP := 0;

        while Assigned(DD) do begin
          {On stocke les taux : Lorsque l'on vient de créer une nouvelle ventilation croisée dans la Tob,
           les variables sont réinitialisées à 0 ou ''. Cependant pour les taux, il faut repartir de la
           valeur du niveau inférieur de la récursivité. Par example :
           Axe 1 : section 11 : 20%, section 12 : 80%
           Axe 2 : section 21 : 10%, section 22 : 50%, section 23 : 40%
           Récursivité Ind = 1 : "Pct1" = 20%
                           Ind = 2 : Sect 11 coisée avec Sect 21 => Pct = "Pct1" * "Pct2" = 2%
                                         à ce moment là, on crée la tob et on réinitialise la variable Pct à 0
                                     Sect 11 coisée avec Sect 22 : à ce moment "Pct1" vaut 0 et "OlP1" 20% =>
                                         réaffectation de "Pct1" à 20% pour calculer Pct = "Pct1" * "Pct2" = 10%
                                     Sect 11 coisée avec Sect 23
                       Ind = 1 : "Pct1" = 80%
                           Ind = 2 : Sect 12 coisée avec Sect 21 ...}

          if Ol1 > 0 then QT1 := Ol1
                     else Ol1 := QT1;
          if Ol2 > 0 then QT2 := Ol2
                     else Ol2 := QT2;
          if Olp > 0 then Pct := OlP
                     else OlP := Pct;

          {Calcul des pourcentages}
          MajVariables(Ind, DD);
          {Récursivité sur l'axe suivant}
          BoucleSur(Ind + 1);
          {on passe à la répartition suivante}
          DD := RetourneTob;
        end;
      end;
    end;

var
  n   : Byte;
  p   : Byte;
  deb : integer ;
begin
  {Initialisation du numéro de ventilation}
  Num := 1;

  {Constitution du tableau de correspondance entre l'ordre des axes paramétrés et leur Axe réel}
  for n := 1 to MaxAxe do begin
    {Initialisation de la valeur du tableau}
    CV[n] := 0;
    {p part de l'axe}
     if n = 1 then  // ajout me pour remplacer iif de ulibwibdows car pb de compile en eaglserver
      deb := 1
       else
        deb := CV[n - 1] + 1;
    for p := deb to MaxAxe do
      if cArrAxe[p] then begin
        CV[n] := p;
        Break;
      end;
  end;

  {Initialisation des "champs" de la table VENTIL}
  InitVariables(0);

  {"Placement" des tobs de manière à ce qu'il n'y ait pas de trou => si T3 <> nil cela signifie que
   T1 <> nil et T2 <> nil}
  T1 := RetourneTobToInd(1);
  T2 := RetourneTobToInd(2);
  T3 := RetourneTobToInd(3);
  T4 := RetourneTobToInd(4);
  T5 := RetourneTobToInd(5);

  D1 := nil;
  D2 := nil;
  D3 := nil;
  D4 := nil;
  D5 := nil;

  if Assigned(T1) then D1 := T1.FindFirst(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[1]), Cpte], True);
  if Assigned(T2) then D2 := T2.FindFirst(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[2]), Cpte], True);
  if Assigned(T3) then D3 := T3.FindFirst(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[3]), Cpte], True);
  if Assigned(T4) then D4 := T4.FindFirst(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[4]), Cpte], True);
  if Assigned(T5) then D5 := T5.FindFirst(['V_NATURE', 'V_COMPTE'], [Nat + IntToStr(CV[5]), Cpte], True);

  {S'il n'y a aucune tob détail pour les sociétés, natures et comptes courants, on sort. Théoriquement cela
   ne devrait pas arrivé !!}
  if not Assigned(D1) and not Assigned(D2) and not Assigned(D3) and not Assigned(D4) and not Assigned(D5) then
    Exit;

  {Lancement de la récursivité}
  BoucleSur(1);
  {Gestion des arrondis}
  GereArrondi(Nat + IntToStr(CV[1]), Cpte, TobN);
end;

{JP 09/11/05 : Si le total des pourcentages d'une ventilation est <> 100, on fait une
               régularisation sur la dernière ligne de la ventilation
{---------------------------------------------------------------------------------------}
procedure TTraitementCroisaxe.GereArrondi(Nat, Cpte : string; var TobN : TOB);
{---------------------------------------------------------------------------------------}
var
  F   : TOB;
  D   : TOB;
  QT1 : Double;
  QT2 : Double;
  Pct : Double;
  Reg : Double;
begin
  QT1 := 0;
  QT2 := 0;
  Pct := 0;
  D := nil;

  F := TobN.FindFirst(['V_NATURE', 'V_COMPTE'], [Nat, Cpte], True);
  while Assigned(F) do begin
    QT1 := QT1 + F.GetDouble('V_TAUXQT1'    );
    QT2 := QT2 + F.GetDouble('V_TAUXQT2'    );
    Pct := Pct + F.GetDouble('V_TAUXMONTANT');
    D := F;
    F := TobN.FindNext(['V_NATURE', 'V_COMPTE'], [Nat, Cpte], True);
  end;

  if Assigned(D) then begin
    {On met l'éventuelle régularisation sur la dernière ligne. Pour les quantités, on teste aussi
     par rapport à 100, car elle ne sont pas nécessairement renseignées}
    Reg := 100 - QT1;
    if (Reg <> 100) and (Reg <> 0) then D.SetDouble('V_TAUXQT1', D.GetDouble('V_TAUXQT1') + Reg);
    Reg := 100 - QT2;
    if (Reg <> 100) and (Reg <> 0) then D.SetDouble('V_TAUXQT2', D.GetDouble('V_TAUXQT2') + Reg);
    Reg := 100 - Pct;
    if Reg <> 0 then D.SetDouble('V_TAUXMONTANT', D.GetDouble('V_TAUXMONTANT') + Reg);
  end;
end;

{Chargement des sections d'attente
{---------------------------------------------------------------------------------------}
procedure TTraitementCroisaxe.SetSectAtt;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT X_AXE, X_SECTIONATTENTE FROM AXE ORDER BY X_AXE', True,-1,'',true);
  try
    while not Q.EOF do begin
           if Q.FindField('X_AXE').AsString = 'A1' then SectAtt1 := Q.FindField('X_SECTIONATTENTE').AsString
      else if Q.FindField('X_AXE').AsString = 'A2' then SectAtt2 := Q.FindField('X_SECTIONATTENTE').AsString
      else if Q.FindField('X_AXE').AsString = 'A3' then SectAtt3 := Q.FindField('X_SECTIONATTENTE').AsString
      else if Q.FindField('X_AXE').AsString = 'A4' then SectAtt4 := Q.FindField('X_SECTIONATTENTE').AsString
      else if Q.FindField('X_AXE').AsString = 'A5' then SectAtt5 := Q.FindField('X_SECTIONATTENTE').AsString;
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;

{JP 07/11/05 : Suppression du paramétrage avant insertion des ventilations croisées
{---------------------------------------------------------------------------------------}
function TTraitementCroisaxe.SupprimeVentiType(Tout : Boolean) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  try
    {Suppression des ventil types}
    ExecuteSQL('DELETE FROM VENTIL');
    {On ne vide la tablette que si l'on décroise}
    if Tout then begin
      ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="VTY"');
      AvertirTable('TTVENTILTYPE');
    end;

    {Pour le moment les clefs de répartition ne sont pas gérées dans le croisement des axes}
    ExecuteSQL('DELETE FROM CLEREPAR');
  except
    on E : Exception do begin
      Result := False;
  {$IFNDEF EAGLSERVER}
      PgiError('Impossible de détruire l''ancien paramétrage :'#13#13 + E.Message);
  {$ENDIF}
    end;
  end;
end;

{JP 02/08/07 : Décroisement des axes : on mémorise les axes qui étaient croisés
{---------------------------------------------------------------------------------------}
procedure TTraitementCroisaxe.MajAxeTraitement(TabAxe: array of Boolean);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 0 to High(TabAxe) do
    cArrAxe[n + 1] := TabAxe[n];
end;

function TTraitementCroisaxe.RechercheJournalOD(vIntAxe : integer):string;
var
  lQJal, lQCodeJal : TQuery;
  lIntCode : integer;
  vTobJal  : TOB;
begin
  result := '';
  vTobJal := nil;
  lIntCode := 0;
  lQCodeJal := OpenSql('SELECT * FROM JOURNAL WHERE J_JOURNAL="' + cTobEcr.Detail[0].GetString('Y_JOURNAL') + '"', True,-1,'',true);
  if not lQCodeJal.Eof then
  begin
    vTobJal := TOB.Create('JOURNAL', nil, -1);
    vTobJal.SelectDB('',lQCodeJal);
  end;

  lQJal := OpenSql('SELECT TOP 1 J_JOURNAL FROM JOURNAL WHERE J_FERME="-" AND J_NATUREJAL="'+ vTobJal.GetString('J_NATUREJAL') +'" AND J_AXE="A' + IntToStr(vIntAxe) + '" ORDER BY J_JOURNAL', True,-1,'',true);
  if not lQJal.eof then
  begin
    result := lQJal.FindField('J_JOURNAL').AsString;
  end
  else
  begin
    if cImport then raise exception.Create('Les journaux pour les OD analytiques doivent être crées');

    begin

      //Génération du code
      while (lIntCode<100) and (ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL="' + 'A'+ Format('%2.2d',[lIntCode]) + '"'))do
      begin
        Inc(lIntcode);
      end;
      if lIntCode = 100 then raise exception.Create('Aucun journal d''od n''a pu être crée. Les journaux A00 à A99 existent déjà.');

      vTobJal.SetString('J_JOURNAL', 'A'+ Format('%2.2d',[lIntCode]));
      vTobJal.setString('J_LIBELLE','Journal OD ' + IntToStr(lIntcode));
      vTobJal.SetString('J_LIBELLE','Journal OD ' + IntToStr(lIntcode));
      vTobJal.SetString('J_AXE', 'A' + IntToStr(vIntAxe));
      //Date
      vTobJal.SetDateTime('J_DATECREATION', Date);
      vTobJal.SetDateTime('J_DATEOUVERTURE', Date);
      vTobJal.SetDateTime('J_DATEMODIF', Date);
      vTobJal.SetDateTime('J_DATEDERNMVT', iDate1900);
      vTobJal.SetDouble('J_DEBITDERNMVT', 0.0);
      vTobJal.SetDouble('J_CREDITDERNMVT', 0.0);
      vTobJal.SetInteger('J_NUMDERNMVT', 0);
      vTobJal.SetDouble('J_TOTALDEBIT', 0.0);
      vTobJal.SetDouble('J_TOTALCREDIT', 0.0);
      vTobJal.SetDouble('J_TOTDEBP', 0.0);
      vTobJal.SetDouble('J_TOTCREP', 0.0);
      vTobJal.SetDouble('J_TOTDEBE', 0.0);
      vTobJal.SetDouble('J_TOTCREE', 0.0);
      vTobJal.SetDouble('J_TOTDEBS', 0.0);
      vTobJal.SetDouble('J_TOTCRES', 0.0);
			// CEGID V9
      vTobJal.SetString('J_CONTREPARTIEAUX', '');
      vTobJal.SetString('J_INVISIBLE', '-');
      vTobJal.SetString('J_LIBELLEAUTO', '');
      vTobJal.SetString('J_TVACTRL', '-');
      // ----

      vTobJal.InsertDB(nil);
      result := vTobJal.GetString('J_JOURNAL');

    end;
    Ferme(lQCodeJal);
  end;
  Ferme(lQJal);
  vTobJal.Free;
end;

procedure TTraitementCroisaxe.AjoutSoldeSection(vTob : TOB);
var
  lIntCpteur : integer;

  procedure AjoutSection(vTobAna : TOB; vStrSection : string; vStrAxe : string; vStrExercice : string);
  var
    lTobTmp : TOB;
    lTobFille : TOB;
  begin
    lTobTmp := cSoldeSection.FindFirst(['SECTION','AXE', 'EXERCICE'], [vStrSection,vStrAxe, vStrExercice], True);
    if lTobTmp = nil then
    begin
      lTobFille := TOB.Create('SOLDESECTION', cSoldeSection, -1);
      lTobFille.AddChampSupValeur('AXE', vStrAxe);
      lTobFille.AddChampSupValeur('SECTION', vStrSection);
      lTobFille.AddChampSupValeur('EXERCICE', vStrExercice);
      lTobFille.AddChampSupValeur('DEBITDERNMVT', vTobAna.GetDouble('Y_DEBIT'));
      lTobFille.AddChampSupValeur('CREDITDERNMVT', vTobAna.GetDouble('Y_CREDIT'));
      lTobFille.AddChampSupValeur('TOTALDEBITDERNMVT', vTobAna.GetDouble('Y_DEBIT'));
      lTobFille.AddChampSupValeur('TOTALCREDITDERNMVT', vTobAna.GetDouble('Y_CREDIT'));
      lTobFille.AddChampSupValeur('DATEDERNMVT', vTobAna.GetDateTime('Y_DATECOMPTABLE'));
      lTobFille.AddChampSupValeur('NUMDERNMVT', vTobAna.GetInteger('Y_NUMEROPIECE'));
      lTobFille.AddChampSupValeur('LIGNEDERNMVT', vTobAna.GetInteger('Y_NUMLIGNE'));
    end
    else
    begin
      lTobTmp.SetInteger('NUMDERNMVT', vTobAna.GetInteger('Y_NUMEROPIECE'));
      lTobTmp.SetInteger('LIGNEDERNMVT', vTobAna.GetInteger('Y_NUMLIGNE'));
      lTobTmp.SetDateTime('DATEDERNMVT', vTobAna.GetDateTime('Y_DATECOMPTABLE'));
      lTobTmp.SetDouble('DEBITDERNMVT', vTobAna.GetDouble('Y_DEBIT'));
      lTobTmp.SetDouble('CREDITDERNMVT', vTobAna.GetDouble('Y_CREDIT'));
      lTobTmp.SetDouble('TOTALDEBITDERNMVT', vTobAna.GetDouble('Y_DEBIT') + lTobTmp.Getdouble('TOTALDEBITDERNMVT'));
      lTobTmp.SetDouble('TOTALCREDITDERNMVT', vTobAna.GetDouble('Y_CREDIT') + lTobTmp.Getdouble('TOTALCREDITDERNMVT'));
    end;
  end;


begin

  if cBooAnaCroisaxe then
  begin
    for lIntCpteur := 1 to 5 do
    begin
      if vTob.GetString('Y_SOUSPLAN' + IntToStr(lIntCpteur)) = '' then Continue;
      AjoutSection(vTob, vTob.GetString('Y_SOUSPLAN' + IntToStr(lIntCpteur)), 'A' + IntToStr(lIntCpteur), vTob.GetString('Y_EXERCICE'));
    end;
  end
  else
  begin
    AjoutSection(vTob, vTob.GetString('Y_SECTION'), vTob.GetString('Y_AXE'), vTob.GetString('Y_EXERCICE'));
  end;
end;



procedure TTraitementCroisaxe.AjoutSoldeJournal(vTob : TOB; vDblDeb, vDblCre : double);
var
    lTobTmp   : TOB;
    lTobFille : TOB;
begin


  lTobTmp := cSoldeJournal.FindFirst(['JOURNAL'], [vTob.GetString('Y_JOURNAL')], True);
  if lTobTmp = nil then
  begin
    lTobFille := TOB.Create('SOLDEJOURNAL', cSoldeJournal, -1);
    lTobFille.AddChampSupValeur('JOURNAL', vTob.GetString('Y_JOURNAL'));
    lTobFille.AddChampSupValeur('DEBITDERNMVT', vDblDeb);
    lTobFille.AddChampSupValeur('CREDITDERNMVT', vDblCre);
    lTobFille.AddChampSupValeur('TOTALDEBITDERNMVT', vDblDeb);
    lTobFille.AddChampSupValeur('TOTALCREDITDERNMVT', vDblCre);
    lTobFille.AddChampSupValeur('DATEDERNMVT', vTob.GetDateTime('Y_DATECOMPTABLE'));
    lTobFille.AddChampSupValeur('NUMDERNMVT', vTob.GetInteger('Y_NUMEROPIECE'));
    lTobFille.AddChampSupValeur('LIGNEDERNMVT', vTob.GetInteger('Y_NUMLIGNE'));
    lTobFille.AddChampSupValeur('EXERCICE', vTob.GetInteger('Y_EXERCICE'));

  end
  else
  begin
    lTobTmp.SetInteger('NUMDERNMVT', vTob.GetInteger('Y_NUMEROPIECE'));
    lTobTmp.SetInteger('LIGNEDERNMVT', vTob.GetInteger('Y_NUMLIGNE'));
    lTobTmp.SetDateTime('DATEDERNMVT', vTob.GetDateTime('Y_DATECOMPTABLE'));
    lTobTmp.SetDouble('DEBITDERNMVT', vTob.GetDouble('Y_DEBIT'));
    lTobTmp.SetDouble('CREDITDERNMVT', vTob.GetDouble('Y_CREDIT'));
    lTobTmp.SetDouble('TOTALDEBITDERNMVT', vDblDeb + lTobTmp.Getdouble('TOTALDEBITDERNMVT'));
    lTobTmp.SetDouble('TOTALCREDITDERNMVT', vDblCre + lTobTmp.Getdouble('TOTALCREDITDERNMVT'));
  end;
end;


procedure TTraitementCroisaxe.MajSoldeSection;
var
  lIntCpteur              : integer;
  XD, XC                  : double ;
  totXD, totXC            : double ;
  StE, lSQL, Ax, Section  : String ;
  Num, Lig                : integer ;
  DateC                   : TDateTime ;
  DE, CE, DP, CP, DS, CS  : Double ;
begin
  //On repart de 0
  lSQL:='UPDATE SECTION SET S_DEBITDERNMVT='+StrfPoint(0.0)+',  S_CREDITDERNMVT='+StrfPoint(0.0)+', '
      +'S_DATEDERNMVT="'+UsDateTime(iDate1900)+'", S_NUMDERNMVT='+IntToStr(0)+', S_LIGNEDERNMVT='+IntToStr(0)+', '
      +'S_TOTALDEBIT='+StrfPoint(0.0)+', S_TOTALCREDIT='+StrfPoint(0.0)+', '
      +'S_TOTDEBE='+StrfPoint(0.0)+', S_TOTCREE='+StrfPoint(0.0)+', '
      +'S_TOTDEBS='+Strfpoint(0.0)+', S_TOTCRES='+StrfPoint(0.0)+', '
      +'S_TOTDEBP='+StrfPoint(0.0)+', S_TOTCREP='+StrfPoint(0.0);

  ExecuteSql(lSQL);

  if not assigned(cSoldeSection) then Exit;
  
  //Maj Solde
  for lIntCpteur := 0 to cSoldeSection.Detail.Count - 1 do
  begin
    XD := cSoldeSection.Detail[lIntCpteur].GetValue('DEBITDERNMVT');
    XC := cSoldeSection.Detail[lIntCpteur].GetValue('CREDITDERNMVT');
    Num := cSoldeSection.Detail[lIntCpteur].GetValue('NUMDERNMVT');
    Lig := cSoldeSection.Detail[lIntCpteur].GetValue('LIGNEDERNMVT') ;
    Ax := cSoldeSection.Detail[lIntCpteur].GetValue('AXE') ;
    Section := cSoldeSection.Detail[lIntCpteur].GetValue('SECTION') ;
    DateC := cSoldeSection.Detail[lIntCpteur].GetValue('DATEDERNMVT') ;
    totXD := cSoldeSection.Detail[lIntCpteur].GetValue('TOTALDEBITDERNMVT');
    totXC := cSoldeSection.Detail[lIntCpteur].GetValue('TOTALCREDITDERNMVT');

    StE := cSoldeSection.Detail[lIntCpteur].GetValue('EXERCICE');

    DE := 0 ; CE := 0 ; DP := 0 ; CP := 0 ; DS := 0 ; CS := 0 ;
    if StE=GetEncours.Code then begin DE := totXD ; CE := totXC; end else
    if StE=GetSuivant.Code then begin DS:=totXD ; CS:=totXC ; end else BEGIN DP := totXC ; CP := totXC ; end ;


    lSQL:='UPDATE SECTION SET S_DEBITDERNMVT='+StrfPoint(XD)+',  S_CREDITDERNMVT='+StrfPoint(XC)+', '
      +'S_DATEDERNMVT="'+UsDateTime(DateC)+'", S_NUMDERNMVT='+IntToStr(Num)+', S_LIGNEDERNMVT='+IntToStr(Lig)+', '
      +'S_TOTALDEBIT=S_TOTALDEBIT+'+StrfPoint(totXD)+', S_TOTALCREDIT=S_TOTALCREDIT+'+StrfPoint(totXC)+', '
      +'S_TOTDEBE=S_TOTDEBE+'+StrfPoint(DE)+', S_TOTCREE=S_TOTCREE+'+StrfPoint(CE)+', '
      +'S_TOTDEBS=S_TOTDEBS+'+Strfpoint(DS)+', S_TOTCRES=S_TOTCRES+'+StrfPoint(CS)+', '
      +'S_TOTDEBP=S_TOTDEBP+'+StrfPoint(DP)+', S_TOTCREP=S_TOTCREP+'+StrfPoint(CP)+' WHERE S_AXE="'+ AX +'" AND S_SECTION="'+ Section +'"' ;

    ExecuteSql(lSQL);
  end;

  //RAZ solde journaux
  if (cSoldeJournal = nil) or (cSoldeJournal.Detail.Count = 0) then Exit;

  lSQL:='UPDATE JOURNAL SET J_DEBITDERNMVT='+StrfPoint(0.0)+', J_CREDITDERNMVT='+StrfPoint(0.0)+', '
      +'J_DATEDERNMVT="'+USDateTime(iDate1900)+'", J_NUMDERNMVT='+IntToStr(0)+', J_TOTALDEBIT=J_TOTALDEBIT+'+Strfpoint(0.0)+', '
      +'J_TOTALCREDIT=J_TOTALCREDIT+'+StrfPoint(0.0)+', J_TOTDEBE=J_TOTDEBE+'+StrfPoint(0.0)+', J_TOTCREE=J_TOTCREE+'+StrfPoint(0.0)+','
      +'J_TOTDEBS=J_TOTDEBS+'+StrfPoint(0.0)+', J_TOTCRES=J_TOTCRES+'+StrfPoint(0.0)+', '
      +'J_TOTDEBP=J_TOTDEBP+'+StrfPoint(0.0)+', J_TOTCREP=J_TOTCREP+'+StrfPoint(0.0)+' WHERE J_NATUREJAL="ODA"';
  ExecuteSql(lSQL);

  for lIntCpteur := 0 to cSoldeJournal.Detail.Count - 1 do
  begin
    XD := cSoldeJournal.Detail[lIntCpteur].GetValue('DEBITDERNMVT');
    XC := cSoldeJournal.Detail[lIntCpteur].GetValue('CREDITDERNMVT');
    Num := cSoldeJournal.Detail[lIntCpteur].GetValue('NUMDERNMVT');
    Section := cSoldeJournal.Detail[lIntCpteur].GetValue('JOURNAL') ;
    DateC := cSoldeJournal.Detail[lIntCpteur].GetValue('DATEDERNMVT') ;
    totXD := cSoldeJournal.Detail[lIntCpteur].GetValue('TOTALDEBITDERNMVT');
    totXC := cSoldeJournal.Detail[lIntCpteur].GetValue('TOTALCREDITDERNMVT');
    StE := cSoldeJournal.Detail[lIntCpteur].GetValue('EXERCICE');

    DE := 0 ; CE := 0 ; DP := 0 ; CP := 0 ; DS := 0 ; CS := 0 ;
    if StE=GetEncours.Code then begin DE := totXD ; CE := totXC; end else
    if StE=GetSuivant.Code then begin DS:=totXD ; CS:=totXC ; end else BEGIN DP := totXC ; CP := totXC ; end ;

    lSQL:='UPDATE JOURNAL SET J_DEBITDERNMVT='+StrfPoint(XD)+', J_CREDITDERNMVT='+StrfPoint(XC)+', '
      +'J_DATEDERNMVT="'+USDateTime(DateC)+'", J_NUMDERNMVT='+IntToStr(Num)+', J_TOTALDEBIT=J_TOTALDEBIT+'+Strfpoint(totXD)+', '
      +'J_TOTALCREDIT=J_TOTALCREDIT+'+StrfPoint(totXC)+', J_TOTDEBE=J_TOTDEBE+'+StrfPoint(DE)+', J_TOTCREE=J_TOTCREE+'+StrfPoint(CE)+','
      +'J_TOTDEBS=J_TOTDEBS+'+StrfPoint(DS)+', J_TOTCRES=J_TOTCRES+'+StrfPoint(CS)+', '
      +'J_TOTDEBP=J_TOTDEBP+'+StrfPoint(DP)+', J_TOTCREP=J_TOTCREP+'+StrfPoint(CP)+' WHERE J_JOURNAL="'+Section+'"' ;

    ExecuteSQL(lSQL);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 13/10/2005
Modifié le ... :   /  /    
Description .. : Synchronise les ventilations de l'écriture passée en
Suite ........ : paramètre pour le champ / valeur passés en paramètre.
Suite ........ : 
Suite ........ : Si vBoForce est à False, la procédure vérifiera avant la 
Suite ........ : présence du champ passé en paramètre dans le tableau 
Suite ........ : référence _recChampsSynchVentil contenant la liste des 
Suite ........ : champs communs stockés Ecriture / Analytique
Mots clefs ... : 
*****************************************************************}
Procedure CReporteVentil ( vTobEcr : TOB ; vChamps : String ; vValeur : Variant ; vBoForce : Boolean ) ;
var lBoATraiter : Boolean ;
    lChampsV    : String ;
    i           : Integer ;

    procedure _traiteChampValeur( vChp : String ; vVal : Variant ) ;
    var lInAxe      : Integer ;
        lInVentil   : Integer ;
    begin
      for lInAxe := 1 to vTobEcr.Detail.Count do
        for lInVentil := 1 to vTobEcr.Detail[ lInAxe - 1 ].Detail.Count do
          vTobEcr.Detail[ lInAxe - 1 ].Detail[ lInVentil - 1 ].PutValue( vChp, vVal ) ;
    end ;

begin

  if vTobEcr = nil then Exit ;
  if vTobEcr.GetValue('E_ANA') <> 'X' then Exit ;

  // champs ok ?
  lBoATraiter := vBoForce ;
  if not lBoATraiter then
    for i := 1 to _InMaxChampsSynchVentil do
      if _recChampsSynchVentil[i] = vChamps then
        begin
        lBoATraiter := True ;
        Break ;
        end ;
  if not lBoATraiter then Exit ;

  // --------------------------------------
  // --- Détermination du nom du champs ---
  // --------------------------------------

  // --> Gestion des qualités
  if (vChamps = 'E_QUALIFQTE1')      then
    lChampsV := 'Y_QUALIFECRQTE1'
  else if (vChamps = 'E_QUALIFQTE2') then
    lChampsV := 'Y_QUALIFECRQTE2'
  else if (vChamps = 'E_QTE1')       then
    lChampsV := 'Y_TOTALQTE1'
  else if (vChamps = 'E_QTE2')       then
    lChampsV := 'Y_TOTALQTE2'

  // --> Report des montants de l'écriture
  else if (vChamps = 'E_DEBIT') or (vChamps = 'E_DEBITDEV') or (vChamps = 'E_CREDIT') or (vChamps = 'E_CREDITDEV') then
    begin
    _traiteChampValeur( 'Y_TOTALECRITURE', vTobEcr.GetDouble('E_CREDIT') + vTobEcr.GetDouble('E_DEBIT')      ) ;
    _traiteChampValeur( 'Y_TOTALDEVISE' , vTobEcr.GetDouble('E_CREDITDEV') + vTobEcr.GetDouble('E_DEBITDEV') ) ;
    Exit ;
    end 

  // --> Champ similaire...
  else lChampsV := FindEtReplace( vChamps, 'E_', 'Y_', False ) ;

  // Affectation de la valeur à toutes les ventilations de la ligne
  if lBoATraiter then
    _TraiteChampValeur( lChampsV, vValeur ) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 13/10/2005
Modifié le ... : 13/10/2005
Description .. : Synchronise les ventilations de l'écriture passée en
Suite ........ : paramètre en prenant pour base la liste des champs
Suite ........ : communs stockés dans le tableau _recChampsSynchVentil
Mots clefs ... :
*****************************************************************}
Procedure CSynchroVentil ( vTobEcr : TOB ) ;
var lStNomChp : String ;
    i         : Integer ;
begin

  if vTobEcr = nil then Exit ;

  for i := 1 to _InMaxChampsSynchVentil do
    begin
    lStNomChp := _recChampsSynchVentil[i] ;
    if vTobEcr.IsFieldModified( lStNomChp ) then
      CReporteVentil( vTobEcr , lStNomChp , vTobEcr.GetValue(lStNomChp), True) ;
    end ;

end ;


function GetFirstAxeAna(Axe: String): Integer;
var
  i: Integer;
begin
  Result := 1;
  if GetAnaCroisaxe then
    for i := 1 to MaxAxe do
      begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) then
        begin
        Result := i;
        break;
        end;
      end
  else if Length(Axe) > 0 then
    Result := StrToInt(Axe[Length(Axe)]);
end;



{ TRestrictionAnalytique }

constructor TRestrictionAnalytique.Create;
begin
  inherited;
  FCompteGene := '';
  FillChar(FModele, sizeof(FModele), #0);
  FillChar(FExclu,  sizeof(FExclu),  #0);
  FAxeCroise := GetAnaCroisaxe;
end;

destructor TRestrictionAnalytique.Destroy;
begin
  inherited;
end;

function TRestrictionAnalytique.GetClause(Axe, TableAna: String; CompteAna: array of String): String;

  function doClause(iAxe: Integer): String;
  begin
    if FChampAna[iAxe]='"'+GetInfoCpta(AxeToFb('A' + IntToStr(iAxe))).Attente+'"' then
      Result := ''
    else
      Result :=
      '   and (   ('+FChampAna[iAxe]+'>=CAD_DEBAXE'+IntToStr(iAxe)+' and '+FChampAna[iAxe]+'<=CAD_FINAXE'+IntToStr(iAxe)+')'+
             ' or ('+FChampAna[iAxe]+' like CAD_DEBAXE'+IntToStr(iAxe)+')'+
           '  )';    
  end;
var
  i:             Integer;
  iAxe:          Integer;
  Prefix:        String;
begin
  Result := '';
  iAxe   := StrToInt(Axe[Length(Axe)]);
  Prefix := TableToPrefixe(TableAna);

  FillChar(FChampAna,  sizeof(FChampAna),  #0);
  for i := 1 to MaxAxe do
    begin
    if (TableAna = 'VENTIL') or (TableAna = 'ANALYTIQ') then
      begin
      if FAxeCroise and (CompteAna[i-1] <> '') then
        FChampAna[i] := CompteAna[i-1]
      else if FAxeCroise then
        FChampAna[i] := Prefix+'_SOUSPLAN'+IntToStr(i)
      else
        FChampAna[i] := Prefix+'_SECTION'
      end
    else if (TableAna = 'SECTION') or (TableAna = 'CSECTION') or(TableAna = 'ANAGUI') then
      begin
      if FAxeCroise then
        begin
        if i = iAxe then
          FChampAna[i] := Prefix+'_SECTION'
        else
          FChampAna[i] := CompteAna[i-1];
        end
      else
        FChampAna[i] := Prefix+'_SECTION'
      end
    end;

  if FAxeCroise then
    begin
    for i := 1 to MaxAxe do
      begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) and (FChampAna[i]<>'') then
        Result := Result + doClause(i);
      end;
    end
  else
    Result := Result + doClause(iAxe);
end;

function TRestrictionAnalytique.GetClauseAttente(iAxe: Integer): String;
  function doClause(iAxe: Integer; var First: Boolean): String;
  begin
    Result := '';
    if Pos('"', FChampAna[iAxe]) = 0 then  {Filtre uniquement sur les champs SQL et pas sur des valeurs}
      begin
      if not First then
        Result := Result+' and';
        Result := Result+' '+FChampAna[iAxe]+'="'+GetInfoCpta(AxeToFb('A' + IntToStr(iAxe))).Attente+'"';
      end;
    if Result<>'' then
      First := False;
  end;
var
  i:     Integer;
  First: Boolean;
begin
  Result := '';
  First  := True;
  if FAxeCroise then
    begin
    for i := 1 to MaxAxe do
      begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) and (FChampAna[i]<>'') then
        Result := Result + doClause(i, First);
      end;
    end
  else
    Result := Result + doClause(iAxe, First);
end;

function TRestrictionAnalytique.GetClauseCompteAutorise(CompteGene, Axe, TableAna: String; CompteAna: array of String): String;
var
  iAxe:      Integer;
  iFirstAxe: Integer;
begin
  iAxe      := StrToInt(Axe[Length(Axe)]);
  iFirstAxe := GetFirstAxe(Axe);
  InitCompteGene(CompteGene);

  Result := ' (1=1)';
  Axe := Trim(Axe);
  if (Axe <> '') and (FModele[iFirstAxe] <> '') then
    begin
    if FExclu[iFirstAxe] then
      Result := ' (not exists('
    else
      Result := ' (exists(';
    Result := Result +
      'select CRA_CODE from CMODELRESTANA'+
        ' join CMODELRESTANADET on (CRA_AXE=CAD_AXE and CRA_CODE=CAD_CODE)' +
        ' where CRA_AXE  = "A'+IntToStr(iFirstAxe)+'"'+
        '   and CRA_CODE = "'+FModele[iFirstAxe]+'"' +
        GetClause(Axe, TableAna, CompteAna);
    Result := Result + ') or ('+GetClauseAttente(iAxe)+') )';
    end;
end;

function TRestrictionAnalytique.GetClauseCompteInterdit(Axe, TableAna: String): String;
var
  iAxe:      Integer;
  CompteAna: array[1..5] of String;
begin
  FillChar(CompteAna, sizeof(CompteAna), #0);
  iAxe   := StrToInt(Axe[Length(Axe)]);
  Result :=
    ' (   (    (CRA_EXCLU="-")'+
    '      and not exists(select CAD_CODE from CMODELRESTANADET'+
    '           where CRA_AXE=CAD_AXE and CRA_CODE=CAD_CODE'+
    '             '+GetClause(Axe, TableAna, CompteAna)+
    '           )'+
    '     )'+
    '  or (    (CRA_EXCLU="X")'+
    '      and exists(select CAD_CODE from CMODELRESTANADET'+
    '           where CRA_AXE=CAD_AXE and CRA_CODE=CAD_CODE'+
    '             '+GetClause(Axe, TableAna, CompteAna)+
    '           )'+
    '     )'+
    ' )'+
    ' and not ('+GetClauseAttente(iAxe)+')';    {Autorise les comptes d'attente}
end;

function TRestrictionAnalytique.GetClauseModeleCompteInterdit(Modele, Axe, TableAna: String): String;
var
  iAxe:  Integer;
begin
  iAxe   := GetFirstAxe(Axe);
  Result := ' exists( select CRA_EXCLU from CMODELRESTANA'+
            '   where CRA_AXE="A'+IntToStr(iAxe)+'" and CRA_CODE="'+Modele+'"' +
            '     and '+GetClauseCompteInterdit(Axe, TableAna) +
            ' )';
end;

function TRestrictionAnalytique.GetFirstAxe(Axe: String): Integer;
var
  i: Integer;
begin
  Result := 1;
  if FAxeCroise then
    for i := 1 to MaxAxe do
      begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) then
        begin
        Result := i;
        break;
        end;
      end
  else if Length(Axe) > 0 then
    Result := StrToInt(Axe[Length(Axe)]);
end;

function TRestrictionAnalytique.GetFromCompteInterdit(Axe, ChampCompteGene, ChampCompteAna: String): String;
var
  NumAxe: String;
begin
  NumAxe := IntToStr(GetFirstAxe(Axe));

  Result :=
    ' join CLIENGENEMODELA on (CLA_GENERAL = '+ChampCompteGene+')'+
    ' join CMODELRESTANA on ('+ChampCompteAna+'=CRA_AXE and CLA_CODE'+NumAxe+'=CRA_CODE)';
end;

procedure TRestrictionAnalytique.InitCompteGene(CompteGene: String);
var
  i:   Integer;
  SQL: String;
  Q:   TQuery;
begin
  if FCompteGene = CompteGene then
    Exit;
  FCompteGene := Trim(CompteGene);
  FillChar(FModele, sizeof(FModele), #0);
  FillChar(FExclu,  sizeof(FExclu),  #0);
  if FCompteGene <> '' then
    begin
    SQL := 'SELECT CLA_GENERAL';
    for i:=1 to MaxAxe do
      begin
      SQL := SQL + ', CLA_CODE'+IntToStr(i);
      SQL := SQL + ', CMODELRESTANA'+IntToStr(i)+'.CRA_EXCLU CRA_EXCLU'+IntToStr(i);
      end;
    SQL := SQL + ' FROM CLIENGENEMODELA';
    for i:=1 to MaxAxe do
      SQL := SQL + ' LEFT JOIN CMODELRESTANA CMODELRESTANA'+IntToStr(i)+
          ' ON (CMODELRESTANA'+IntToStr(i)+'.CRA_AXE="A'+IntToStr(i)+'"'+
            ' and CLA_CODE'+IntToStr(i)+'=CMODELRESTANA'+IntToStr(i)+'.CRA_CODE)';
    SQL := SQL + ' WHERE CLA_GENERAL="'+FCompteGene+'"';
    Q := OpenSQL(SQL, True,-1,'',true);
    try
      if not Q.EOF then
        begin
        for i:=1 to MaxAxe do
          begin
          FModele[i] := Q.FindField('CLA_CODE'+IntToStr(i)).AsString;
          FExclu[i]  := Q.FindField('CRA_EXCLU'+IntToStr(i)).AsString='X';
          end;
        end;
    finally
      Ferme( Q ) ; // SBO 07/05/2006 ( ne pas hésiter à tester en mode SAV ;-) )
      FreeAndNil(Q);
      end;
    end;
end;

function TRestrictionAnalytique.IsCompteGeneAutorise(CompteGene, Axe, Nature, Modele: String): Boolean;
var
  CompteAna:  array[1..MaxAxe] of String;
  SQL:        String;
begin
  FillChar(CompteAna, sizeof(CompteAna), #0);

  SQL :=
      ' SELECT *' +
      ' FROM VENTIL ' + GetFromCompteInterdit(Axe, '"'+CompteGene+'"', '"'+Axe+'"')+
      ' WHERE V_NATURE="'+Nature+'"'+
      '   AND V_COMPTE="'+Modele+'"'+
      '   AND ' + GetClauseCompteInterdit(Axe, 'VENTIL');
  Result := not ExisteSQL(SQL);
end;

function TRestrictionAnalytique.IsModelExclu(CompteGene, Axe: String): Boolean;
var
  iFirstAxe:     Integer;
begin
  {b FP 19/04/2006 FQ17728}
  iFirstAxe   := GetFirstAxe(Axe);
  InitCompteGene(CompteGene);
  Result := (FChampAna[iFirstAxe] <> '') and FExclu[iFirstAxe];
  {e FP 19/04/2006}
end;

procedure TRestrictionAnalytique.SetAxeCroise(const Value: Boolean);
begin
  FAxeCroise := Value;
end;

function TRestrictionAnalytique.VerifModelVentil(CompteGene: String; Q: TQuery; Table, Axe: String; QrySansFiltre: String; ForceMsgErr: Boolean = False): Boolean;
var
  MsgErr:        String;
  iFirstAxe:     Integer;        {FP 19/04/2006 FQ17725}
  QryAll:        TQuery;
begin
  Result      := True;
  MsgErr      := 'Pour l''axe '+Axe+','+Chr($D)+Chr($A)+
    ' il existe des incompatibilités entre le modèle de ventilations à appliquer et les modèles de restrictions analytiques attachés aux comptes généraux sélectionnés.';

  {b FP 19/04/2006 FQ17725}
  iFirstAxe   := GetFirstAxe(Axe);
  InitCompteGene(CompteGene);

  if (QrySansFiltre <> '') and
     (FChampAna[iFirstAxe] <> '') then  {Vérifie la présence d'un modèle de restriction}
    begin
    QryAll := OpenSQL(QrySansFiltre, True,-1,'',true);
    {Retourne False seulement si le nb d'elts de la query sans restriction <> du nb d'elts de la query avec restriction}
    if (QryAll.RecordCount <> 0) and (QryAll.RecordCount <> Q.RecordCount) then
      Result := False;
    Ferme(QryAll);
    FreeAndNil(QryAll);
    end;
  {e FP 19/04/2006}

  {$IFNDEF EAGLSERVER}
  if (not Result) then
    PGIInfo(TraduireMemoire(MsgErr));
  {$ENDIF}
end;

{b FP 19/04/2006 FQ17725+FQ17831}
function TRestrictionAnalytique.VerifModelVentil(CompteGene: String; Q, QAll: TOB; Table, Axe: String; ForceMsgErr: Boolean = False): Boolean;
var
  MsgErr:        String;
  iFirstAxe:     Integer;
begin
  Result      := True;
  MsgErr      := 'Pour l''axe '+Axe+','+Chr($D)+Chr($A)+
    ' il existe des incompatibilités entre la ventilation à appliquer et le modèle de restriction rattaché au compte général.';

  iFirstAxe   := GetFirstAxe(Axe);
  InitCompteGene(CompteGene);

  if (QAll.Detail.Count > 0) and
     (FChampAna[iFirstAxe] <> '') then  {Vérifie la présence d'un modèle de restriction}
    begin
    {Retourne False seulement si le nb d'elts de la query sans restriction <> du nb d'elts de la query avec restriction}
    if QAll.Detail.Count <> Q.Detail.Count then
      Result := False;
    end;

  {$IFNDEF EAGLSERVER}
  if (not Result) then
    PGIInfo(TraduireMemoire(MsgErr));
  {$ENDIF}
end;
{e FP 19/04/2006}

// ===============================================
// === Fonctions de validation de l'analytique ===
// ===============================================

{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}

procedure _CTestErreur( var vInErreur : integer ; vInfo : TInfoEcriture ) ;
var
 j : integer ;
begin
 for j := low(vInfo.ErreurIgnoree) to high(vInfo.ErreurIgnoree) do
  if vInErreur = vInfo.ErreurIgnoree[j] then
   begin
    vInErreur := RC_PASERREUR ;
    break ;
   end ;
end ;

procedure _CTestRecErreur( var vErreur : TRecError ; vInfo : TInfoEcriture ) ;
var
 j : integer ;
begin
 for j := low(vInfo.ErreurIgnoree) to high(vInfo.ErreurIgnoree) do
  if vErreur.RC_Error = vInfo.ErreurIgnoree[j] then
   begin
    vErreur.RC_Error := RC_PASERREUR ;
    break ;
   end ;
end ;


function CIsValidSection  ( vTobLigne : Tob ; vInfo : TInfoEcriture ) : integer ;
var lStSection : string ;
    lStAxe     : string ;
begin

 result          := RC_PASERREUR ;

 lStSection      := vTobLigne.GetString('Y_SECTION') ;
 lStAxe          := vTobLigne.GetString('Y_AXE') ;
 if vInfo.LoadSection( lStSection, lStAxe ) then
  lStSection      := vInfo.StSection ;

 if ( trim(lStSection) = '' ) then
   begin
   result := RC_YSECTIONOBLIG ;
   _CTestErreur( Result, vInfo ) ;
   end;

  if result = RC_PASERREUR then
    if ( vInfo.Section.InIndex = - 1 ) then
     begin
     result := RC_YSECTIONINCONNUE ;
     _CTestErreur( Result, vInfo ) ;
     end;

  if result = RC_PASERREUR then
    if ( vInfo.Section_GetValue('S_FERME') = 'X' ) then
      begin
      result :=  RC_YSECTIONFERMEE ;
      _CTestErreur( Result, vInfo ) ;
      end ;

  if result = RC_PASERREUR then
    if vTobLigne.GetString('Y_AXE') <> vInfo.Section.StAxe then
      begin
      result := RC_YAXESECTION ;
      _CTestErreur( Result, vInfo ) ;
      end ;

 vTobLigne.PutValue('Y_SECTION', lStSection ) ;

end ;


function CIsValidAxeBudget  ( vAxe : integer ; vTobAxe : Tob ; vInfo : TInfoEcriture ) : integer ;
var lQBudjal   : TQuery ;
    lStAxeBud  : String ;
    lBoTrouv   : boolean ;
    i          : integer ;
    lStJalBud  : string ;
begin

  result          := RC_PASERREUR ;

  // test la variable info
  if vInfo = nil then
    begin
    {$IFNDEF EAGLSERVER}
    MessageAlerte('La variable Info n''est pas initialisée ! ') ;
    {$ENDIF}
    raise EAbort.Create('') ;
    end;

  // Test de la tob
  if (vTobAxe = nil) then
    begin
    {$IFNDEF EAGLSERVER}
    MessageAlerte('La variable vTob n''est pas initialisée ! ') ;
    {$ENDIF}
    raise EAbort.Create('') ;
    end;

  if vTobAxe.Detail.Count = 0 then Exit ;
  if not vInfo.LoadCompte( vTobAxe.Detail[0].getString('Y_GENERAL') ) then Exit ;

  lStAxeBud     := '' ;
  lBoTrouv      := True ;
{$IFNDEF NOVH}
  lStJalBud     := VH^.JalCtrlBud ;
{$ENDIF NOVH}

  if lStJalBud = '' then Exit ;

  lQBudjal := OpenSQL('SELECT BJ_AXE FROM BUDJAL WHERE BJ_BUDJAL="' + lStJalBud + '"', True,-1,'',true ) ;
  if Not lQBudjal.EOF then
    lStAxeBud := lQBudjal.FindField('BJ_AXE').AsString ;
  Ferme( lQBudjal ) ;

  if ( (lStAxeBud='') or (Length(lStAxeBud)<2) ) then Exit ;
  if ValeurI( lStAxeBud[2] ) <> vAxe then Exit ;

  for i := 0 to vTobAxe.Detail.Count-1 do
    begin
    if not vInfo.LoadSection( vTobAxe.Detail[i].getString('Y_SECTION'), vTobAxe.Detail[i].getString('Y_AXE') ) then continue ;
    lBoTrouv := ExisteSQL('SELECT CX_TYPE FROM CROISCPT WHERE CX_TYPE="GEN" AND CX_JAL="' + lStJalBud +
                               '" AND CX_COMPTE="' + vInfo.StCompte + '" AND CX_SECTION="' + vInfo.StSection + '"' ) ;
    if not lBoTrouv then Break ;
    end ;

  if not lBoTrouv then
    begin
    result := RC_YSECTIONBUDGET ;
    _CTestErreur( Result, vInfo ) ;
    end ;

end ;


function CIsValidAxe ( vAxe : integer ; vTobAxe : Tob ; vInfo : TInfoEcriture ; vBoVerifQte : Boolean = False ; vBoVerifBudget : Boolean = False ) : TRecError ;
var i           : integer ;
    lTobAna     : Tob ;
    lPourc      : double ;
    lMontant    : double ;
    lMontantDev : double;
{    lPourcQte1    : double ;
    lPourcQte2    : double ;
}
begin
  result.RC_Error := RC_PASERREUR ;
  result.RC_Axe   := vAxe ;

  // test la variable info
  if vInfo = nil then
    begin
    {$IFNDEF EAGLSERVER}
    MessageAlerte('La variable Info n''est pas initialisée ! ') ;
    {$ENDIF}
    raise EAbort.Create('') ;
    end;

  // Test de la tob
  if (vTobAxe = nil) then
    begin
    {$IFNDEF EAGLSERVER}
    MessageAlerte('La variable vTob n''est pas initialisée ! ') ;
    {$ENDIF}
    raise EAbort.Create('') ;
    end;

  // Test Au moins une ligne !
  if vTobAxe.Detail.count = 0 then
    begin
    result.RC_Error    := RC_YAXEVIDE ;
    _CTestRecErreur( result, vInfo ) ;
    if result.RC_Error <> RC_PASERREUR then
      Exit ;
    end ;

  lPourc      := 0 ;
  lMontant    := 0 ;
  lMontantDEv := 0 ;

  // Test des lignes
  for i := 0 to vTobAxe.Detail.count - 1 do
    begin
    lTobAna := vTobAxe.Detail[i] ;

    // Test ligne
    result := CIsValidLigneAna( lTobAna, vInfo ) ;
    if result.RC_Error <> RC_PASERREUR then Break ;

    // Totalisations
    lPourc      := lPourc      + lTobAna.GetDouble('Y_POURCENTAGE') ;
    lMontant    := lMontant    + GetMontant( lTobAna ) ;
    lMontantDev := lMontantDev + GetMontantDev( lTobAna ) ;
//    lPourcQte1  := lPourcQte1  + lTobAna.GetDouble('Y_POURCENTQTE1') ;
//    lPourcQte2  := lPourcQte2  + lTobAna.GetDouble('Y_POURCENTQTE2') ;
    end ;

  // test solde poucentage
  if result.RC_Error = RC_PASERREUR then
    if Arrondi( lPourc - 100, ADecimP ) <> 0  then
      begin
      result.RC_Error := RC_YSOLDEPOURC ;
      _CTestRecErreur( result, vInfo ) ;
      end ;

  // test solde montant
  lTobAna := vTobAxe.Detail[0] ;
  if result.RC_Error = RC_PASERREUR then
    if Arrondi( lMontant - lTobAna.GetDouble('Y_TOTALECRITURE') , V_PGI.OkDecV ) <> 0 then
      begin
      result.RC_Error := RC_YSOLDEMONTANT ;
      _CTestRecErreur( result, vInfo ) ;
      end ;

  // test solde montant dev
  if result.RC_Error = RC_PASERREUR then
    if Arrondi( lMontantDev - lTobAna.GetDouble('Y_TOTALDEVISE') , V_PGI.OkDecV ) <> 0 then
      begin
      result.RC_Error := RC_YSOLDEMONTANTDEV ;
      _CTestRecErreur( result, vInfo ) ;
      end ;

  if vBoVerifBudget and (result.RC_Error = RC_PASERREUR) then
    result.RC_Error := CIsValidAxeBudget( vAxe, vTobAxe, vInfo )

{
  if vBoVerifQte and (result.RC_Error = RC_PASERREUR) then
    begin
    // test solde Qte1
    if Arrondi( lPourcQte1 - 100, ADecimP ) <> 0  then
    // RC_YSOLDEQTE1
     // A venir
       ;
    // test solde Qte2
    if result.RC_Error = RC_PASERREUR then
      if Arrondi( lPourcQte2 - 100, ADecimP ) <> 0  then
       // RC_YSOLDEQTE2
       // A venir
       ;
    end ;
}

end ;

function CIsValidMontantAna ( vTobLigne : Tob ; vInfo : TInfoEcriture ) : integer ;
begin

  result := RC_PASERREUR ;

  if vTobLigne.GetDouble('Y_POURCENTAGE') = 0 then
    begin
    result := RC_YPOURCINEXISTANT ;
    _CTestErreur( result, vInfo ) ;
    end ;

  if result = RC_PASERREUR then
    begin
    result := CIsValidMontant( vTobLigne.GetValue('Y_DEBIT') ) ;
    _CTestErreur( result, vInfo ) ;
    end ;

  if result = RC_PASERREUR then
    begin
    result := CIsValidMontant( vTobLigne.GetValue('Y_CREDIT') ) ;
    _CTestErreur( result, vInfo ) ;
    end ;

  if result = RC_PASERREUR then
    if ( vTobLigne.GetValue('Y_DEBIT') = 0 ) and ( vTobLigne.GetValue('Y_CREDIT') = 0 ) then
      begin
      result := RC_YMONTANTINEXISTANT ;
      _CTestErreur( result, vInfo ) ;
      end ;

end ;

function CIsValidLigneAna ( vTobLigne : Tob ; vInfo : TInfoEcriture ) : TRecError ;
begin
  result.RC_Error   := RC_PASERREUR ;
  result.RC_Methode := 'UlibAnalytique.CIsValidLigneAna' ;

  // test la variable info
  if vInfo = nil then
    begin
    result.RC_Message := 'La variable Info n''est pas initialisée ! ' ;
    {$IFNDEF EAGLSERVER}
    MessageAlerte(result.RC_Message) ;
    {$ENDIF}
    raise EAbort.Create('') ;
    end;

 if vTobLigne.NomTable <> 'ANALYTIQ' then
  begin
   result.RC_Message := 'la tob n''est pas du bon type : ' + vTobLigne.NomTable + ' ! ' ;
   {$IFNDEF EAGLSERVER}
   MessageAlerte(result.RC_Message) ;
   {$ENDIF}
   raise EAbort.Create('') ;
  end ;

  result.RC_Valeur := vTobLigne.GetString('Y_SECTION') ;

  // test section
  result.RC_Error := CIsValidSection( vTobLigne, vInfo ) ;

  // pourcentage et montant
  if result.RC_Error = RC_PASERREUR then
   result.RC_Error := CIsValidMontantAna( vTobLigne, vInfo ) ;

end ;

function CIsValidVentil ( vTobEcr : Tob ; vInfo : TInfoEcriture ; vBoVerifQte : Boolean = False ; vBoVerifBudget : Boolean = False ) : TRecError ;
var i      : integer ;
    lStCpt : string ;
begin

  result.RC_Error    := RC_PASERREUR ;
  result.RC_Axe      := 0 ;
  result.RC_Methode  := 'ULibAnalytique.CIsValidVentil' ;

  // test la variable info
  if vInfo = nil then
    begin
    result.RC_Message  := 'La variable Info n''est pas initialisée ! ' ;
    {$IFNDEF EAGLSERVER}
    MessageAlerte(result.RC_Message) ;
    {$ENDIF}
    raise EAbort.Create('') ;
    end;

  // Test de la tob
  if (vTobEcr = nil) or (vTobEcr.Detail.count=0) then
    begin
    result.RC_Message  := 'La variable vTobEcr n''est pas initialisée ! (CIsValidVentil)' ;
    {$IFNDEF EAGLSERVER}
    MessageAlerte(result.RC_Message) ;
    {$ENDIF} 
    raise EAbort.Create('') ;
    end;


  // Récup info compte général
  lStCpt := vTobEcr.GetString('E_GENERAL') ;
  vInfo.LoadCompte( lStCpt ) ;
  if not vInfo.Compte.IsVentilable then Exit ;

  // parcours des axes
  for i := 1 to MAXAXE do
    begin
    // Test Axe ventilable
    if not vInfo.Compte.IsVentilable( i, vInfo.Compte.InIndex ) then Continue ;
    // test ventil de l'axe
    result := CIsValidAxe( i , vTobEcr.detail[ i - 1 ], vInfo, vBoVerifQte, vBoVerifBudget ) ;
    if result.RC_Error <> RC_PASERREUR then Break ;
    end ;

end ;

procedure CGereFenAna ( vStGen, vStNatJal : string ; vBoUnSeulAxe : boolean ; var vBoOuvreAna,vBoGestionTva,vBoVentilTVA : boolean ) ;
var
 lBoJalEnc   : boolean ;
 lBoJalDebit :  boolean ;
begin

 vBoOuvreAna         := true ;
 vBoGestionTva       := false ;
 vBoVentilTVA        := false ;

 if not GetParamSocSecur('SO_CPPCLSAISIETVA',false) then exit ;

 vStGen              := Copy(vStGen,1,2) ;
 lBoJalEnc           := ( vStNatJal = 'BQE' ) or ( vStNatJal = 'CAI' ) or ( vStNatJal = 'OD' ) ;
 lBoJalDebit         := vStNatJal = 'VTE' ;


 if GetParamSocSecur('SO_OUITVAENC',False) then
  begin
   if ( vStGen = '41' ) and  ( lBoJalEnc ) then
    begin
     vBoOuvreAna       := true ;
     vBoGestionTva     := true ;
     vBoVentilTVA      := false ;
    end
     else
      if ( vStGen = '41' ) and vBoUnSeulAxe and not lBoJalEnc then
       begin
        vBoOuvreAna      := false ;
        vBoGestionTva    := false ;
        vBoVentilTVA     := true ;
       end
          else
            if ( vStGen = '41' ) and ( not lBoJalEnc ) then
             begin
              vBoOuvreAna      := true ;
              vBoGestionTva    := false ;
              vBoVentilTVA     := true ;
             end ;
    end
     else
      if ( vStGen = '70' ) and  ( lBoJalDebit ) then
        begin
         vBoOuvreAna       := true ;
         vBoGestionTva     := true ;
         vBoVentilTVA      := false ;
        end
         else
          if ( vStGen = '70' ) and vBoUnSeulAxe and not lBoJalDebit then
           begin
            vBoOuvreAna      := false ;
            vBoGestionTva    := false ;
            vBoVentilTVA     := true ;
           end
              else
                if ( vStGen = '70' ) and ( not lBoJalDebit ) then
                 begin
                  vBoOuvreAna      := true ;
                  vBoGestionTva    := false ;
                  vBoVentilTVA     := true ;
                 end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 25/04/2007
Modifié le ... : 02/05/2007
Description .. : - LG - 25/04/2007 - ajout du parma CodeGuide pour gerer 
Suite ........ : les guide analytiq
Suite ........ : - LG - 02/05/2007 - ajout de la fct AlloueAxe
Mots clefs ... : 
*****************************************************************}
procedure CVentilerTOB( vTOBEcr : TOB ; vInfo : TInfoEcriture ; QueAttente : boolean = false ; vCodeGuide : string = ''  ) ;
var
 lBoOuvreAna    : boolean ;
 lBoGestionTva  : boolean ;
 lBoVentilTVA   : boolean ;
 i              : integer ;
 lstNatJal      : string ;
 lStVentil      : string ;
 lInAxeTva      : integer ;
 lRestriction   : TRestrictionAnalytique ;
begin

 if ( not vInfo.LoadCompte(vTOBEcr.GetString('E_GENERAL')) ) or ( not vInfo.Loadjournal(vTOBEcr.GetString('E_JOURNAL')) ) then exit ;

 lStNatJal   := vInfo.Journal.GetValue('J_NATUREJAL') ;

 if not vInfo.Compte.IsVentilable then exit ;

 if vTOBEcr.Detail.Count < MAXAXE then
  begin
   vTOBEcr.ClearDetail ;
   AlloueAxe( vTOBEcr ) ;
  end ;

 lRestriction := TRestrictionAnalytique.Create ;

 try

 CGereFenAna(vTOBEcr.GetValue('E_GENERAL'),lStNatJal,false,lBoOuvreAna,lBoGestionTva,lBoVentilTVA) ;

 lStVentil := vInfo.Compte.GetValue('G_VENTILABLE1') +
              vInfo.Compte.GetValue('G_VENTILABLE2') +
              vInfo.Compte.GetValue('G_VENTILABLE3') +
              vInfo.Compte.GetValue('G_VENTILABLE4') +
              vInfo.Compte.GetValue('G_VENTILABLE5') ;

 if lBoVentilTVA then
  begin
   lInAxeTva := StrToInt(Copy(GetParamSocSecur('SO_CPPCLAXETVA', '  '),2,1)) ;
   for i:=1 to 5 do
    if ( lStVentil[i] = 'X' ) and  ( lInAxeTva = i ) then
      CVentilerTOBTva(vTOBEcr)
     else
      if ( lStVentil[i] = 'X' ) and ( not QueAttente ) then
       begin
        if not CVentilerVentil ( vTOBEcr,i,V_PGI.OkDecV,'V_NATURE="GE',vTOBEcr.GetString('E_GENERAL'),vInfo.Dossier,false,lRestriction ) then
         CVentilerAttente(vTOBEcr,i) ;
       end ;
   ArrondirAnaTOB(vTOBEcr,V_PGI.OkDecV) ;
  end
   else
    begin
     if vCodeGuide <> '' then
      VentilerTob(vTOBEcr, vCodeGuide , vTOBEcr.GetValue('E_NUMLIGNE'), V_PGI.OkDecV, false)
       else
        VentilerTob(vTOBEcr, '' , 0, V_PGI.OkDecV, false) ;  // si on ne passe pas de guide, on garde l'ancien appel
    end ;

  vTOBEcr.PutValue('E_ANA' , 'X' ) ;   

 finally
  lRestriction.Free ;
 end ;   

end ;


function CAOuvrirVentil  ( vTobEcr   : Tob ; vTobScenario : Tob ; vInfo : TInfoEcriture ; var vNumAxe : integer ; vBoParam : Boolean ) : Boolean ;
var lStCpt       : string ;
    lStAxe       : string ;
    lErr         : TRecError ;
    lBoUnSeulAxe : boolean ;
    lSt          : string ;
    lBoGestionTva : boolean ;
    lBoVentilTva  : boolean ;
    lBoModeBor    : boolean ;
begin

  result  := False ;
  vNumAxe := -1 ;

  // pas de ligne décriture on sort...
  if not Assigned( vTobEcr ) then Exit ;

  lBoModeBor := ( vTobEcr.getValue('E_MODESAISIE') = 'BOR' ) or ( vTobEcr.getValue('E_MODESAISIE') = 'LIB' ) ;

  // Ventilation déjà présente, on sort...
  if VentilationExiste( vTobEcr ) then
    begin
    lErr := CIsValidVentil( vTobEcr, vInfo ) ;
    if lErr.RC_Error <> RC_PASERREUR then
      begin
      vNumAxe := lErr.RC_Axe ;
      result := True ;
      end ;
    Exit ;
    end ;

  if lBoModeBor and ( not GetParamSocSecur('SO_ZGEREANAL',False) ) and not vBoParam then
   begin
    if vTobEcr.Detail.Count < MAXAXE then
     begin
      vTobEcr.ClearDetail ;
      AlloueAxe( vTobEcr ) ;
     end ;
    VentilerTob( vTobEcr, '', 0, V_PGI.OkDecV,false ) ;
    exit ;
   end ;

  lStCpt := vTobEcr.GetString('E_GENERAL') ;

  // pas de compte on sort...
  if ( not vInfo.LoadCompte( lStCpt ) ) or ( not vInfo.LoadJournal(vTobEcr.GetString('E_JOURNAL') ) ) then Exit ;

  lSt     := vInfo.Compte.GetValue('G_VENTILABLE1') +
             vInfo.Compte.GetValue('G_VENTILABLE2') +
             vInfo.Compte.GetValue('G_VENTILABLE3') +
             vInfo.Compte.GetValue('G_VENTILABLE4') +
             vInfo.Compte.GetValue('G_VENTILABLE5') ;

  vNumAxe := Pos('X',lSt) ;

  if vNumAxe = 0 then exit ;

  lBoUnSeulAxe :=  Pos('X',Copy(lSt,vNumAxe+1,5) ) = 0 ;

  // On appelle cette fonction pour savoir si l'on doit gerer la tva et si on doit ouvrir la fenetre de saisie
  CGereFenAna( vTobEcr.getValue('E_GENERAL'),vInfo.Journal.GetValue('J_NATUREJAL'),lBoUnSeulAxe,result,lBoGestionTva,lBoVentilTva) ;

  if lBoModeBor and not vBoParam then
   result := result and GetParamSocSecur('SO_ZSAISIEANAL',False) ;

  if vTobEcr.Detail.Count < MAXAXE then
   begin
    vTobEcr.ClearDetail ;
    AlloueAxe( vTobEcr ) ;
   end ;

  // indicateur de nouvelle saisie (FQ18147 Remise en place de l'annulation)
  if vTobEcr.GetNumChamp('NEWVENTIL') > 0
    then vTobEcr.putValue( 'NEWVENTIL', 'X' )
    else vTobEcr.AddChampSupValeur( 'NEWVENTIL', 'X' ) ;


  // ======================
  // === 1. VENTIL TVA ====
  // ======================
  if lBoVentilTva then
   CVentilerTOB(vTobEcr,vInfo,false)
  // =============================================
  // === 2. MODE BOR et paramSoc non coché... ====
  // =============================================
  else if ( not result ) and lBoModeBor and (not vBoParam) then
      VentilerTob( vTobEcr, '', 0, V_PGI.OkDecV, false )
  // ==================================================================================
  // === 3. PAS DE SCENARIO : on récupère les valeurs par défaut de la page SAISIE ====
  // ==================================================================================
  else if ( vTobScenario = nil ) then
    begin
    if GetParamsocDossierSecur( 'SO_ZSAISIEANAL', false, vInfo.dossier ) then
      VentilerTob( vTobEcr, '', 0, V_PGI.OkDecV, False, vInfo.dossier, False ) // Calcul ventilation par défaut non complète
    else
      begin
      // Calcul ventilation avec solde sur section attente
      VentilerTob( vTobEcr, '', 0, V_PGI.OkDecV, False, vInfo.dossier, True ) ;
      // Pas d'ouverture de la saisie
      result := False ;
      end ;
    end
  // ============================================================
  // === 4. présence d'un scénario avec ouverture analytique ====
  // ============================================================
  else begin
    if ( vTobScenario.GetString('SC_OUVREANAL') = 'X' ) then
      begin
      // Axe Préférenctiel...
      lStAxe := vTobScenario.GetString('SC_NUMAXE') ;
      if ( lStAxe <> '' ) and vInfo.Compte.IsVentilable( ValeurI(lStAxe[2]), vInfo.Compte.InIndex ) then
        vNumAxe := ValeurI(lStAxe[2]) ;
      // Calcul ventilation non complète
      VentilerTob( vTobEcr, '', 0, V_PGI.OkDecV, False, vInfo.dossier, False ) ;
      // Verif ventilation
      lErr := CIsValidVentil( vTobEcr, vInfo ) ;
      if lErr.RC_Error <> RC_PASERREUR then
        vNumAxe := lErr.RC_Axe ;
      end
    else
  // ============================================================
  // === 5. présence d'un scénario sans ouverture analytique ====
  // ============================================================
      begin
      // ==> 3.1 : présence de ligne avec pourcentage à 0 dans la ventilation par défaut...
      if existeSQL( 'SELECT V_COMPTE FROM ' + GetTableDossier( vInfo.Dossier, 'VENTIL')
                    + ' WHERE V_NATURE LIKE "GE%" AND V_COMPTE ="' + vInfo.StCompte
                    + '" AND V_TAUXMONTANT=0' ) then
        begin
        // Calcul ventilation non complète
        VentilerTob( vTobEcr, '', 0, V_PGI.OkDecV, False, vInfo.dossier, False ) ;
        // Verif ventilation
        lErr := CIsValidVentil( vTobEcr, vInfo ) ;
        if lErr.RC_Error <> RC_PASERREUR
          then vNumAxe := lErr.RC_Axe ;
        end
      else
        begin
        // Calcul ventilation avec solde sur section attente
        VentilerTob( vTobEcr, '', 0, V_PGI.OkDecV, False, vInfo.dossier, True ) ;
        // Pas d'ouverture de la saisie
        result := False ;
        end ;
      end ;
    end ;

end ;

{$ENDIF}
{$ENDIF} //CCADM

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 26/03/2007
Modifié le ... : 27/08/2007
Description .. : DEV 5942 : Amélioration analytique
Suite ........ : Détermination auto des paramètres de ventilations du
Suite ........ : comptes en fonction des paramètres des sections :
Suite ........ :   - fourchette de comptes
Suite ........ :   - Unités d'oeuvre
Suite ........ :
Suite ........ : FQ 21259 27/08/2007 : On ne mélange plus le tableau
Suite ........ : d'appel des axes à tester et le tableau résultats des axes 
Suite ........ : ventilés
Mots clefs ... : 
*****************************************************************}
function InitialiseVentilDefaut ( Compte : string ; SectAtt : Array of Boolean ; var AxeVent : Array of Boolean ) : boolean ;
var
    lStAxe    : string ;
    lTobVent  : Tob ;
    lTobL     : Tob ;
    lInNb     : integer ;
    lTotalUO  : double ;
    lVentilUO : double ;
    lInAxe    : integer ;

    lQRub     : TQuery;
    lTSection : TOB;
    lstSQL    : string;
    i         : integer;
    SQL       : string;

    procedure _finaliseVentil ;
    var lPourc : Double ;
        lTotP  : Double ;
        i      : integer ;
      begin
      lTotP := 0.0 ;
      for i := 0 to lTobVent.Detail.count - 1 do
        begin
        lTobL   := lTobVent.Detail[ i ] ;
        // Calcul pourcentage
        if i < lTobVent.Detail.count - 1 then
          begin
          lPourc  := Arrondi( ( lTobL.GetDouble('V_MONTANT') / lTotalUO ) * 100, 4 ) ;
          lTotP   := lTotP + lPourc ;
          end
        else
          begin
          lPourc  := Arrondi( 100.0 - lTotP, 4 ) ;
          lTotP   := 100.0 ;
          end ;
        // MAj objet
        lTobL.PutValue( 'V_TAUXMONTANT', lPourc ) ;
        lTobL.PutValue( 'V_TAUXQTE1',    lPourc ) ;
        lTobL.PutValue( 'V_TAUXQTE2',    lPourc ) ;
        end ;
      end ;

begin
  result := False ;
  if length( Compte ) <> GetInfoCpta (fbGene).Lg then Exit ;

  // init variables
  lTobVent  := TOb.Create('Calcul_Ventil', nil, -1) ;
  lStAxe    := 'A1' ;
  lInAxe    := 1 ;
  lInNb     := 1 ;
  lTotalUO  := 0 ;

  // initialisation du tableau de résultat
  for i := 0 to 4 do
    AxeVent[i]  := false ;

  try  // finally
    try  // Except
      beginTrans ;
      // Requête
      { Nouveau code qui permet de trouver la liste des sections associées au compte via les tranches de comptes }
      lTSection := TOB.Create ('', nil, -1);
      { BVE 01.10.07 : Bug si plusieurs rubrique avec meme code dans des dossiers différents
      lQRub := OpenSQL ('SELECT RB_RUBRIQUE,RB_COMPTE1,RB_EXCLUSION1 FROM RUBRIQUE WHERE RB_CLASSERUB="TRA" ORDER BY RB_AXE',True); }
      SQL := 'SELECT RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1 ' +
             'FROM RUBRIQUE ' +
             'WHERE RB_CLASSERUB="TRA" ' +
             'AND (RB_PREDEFINI<>"DOS" OR (RB_PREDEFINI="DOS" AND RB_NODOSSIER="'+V_PGI.NoDossier+'"))'+
             'ORDER BY RB_RUBRIQUE, RB_NODOSSIER DESC , RB_PREDEFINI DESC, RB_DATEVALIDITE ASC, RB_AXE';
      lQRub := OpenSQL (SQL,True,-1,'',true);
      try
        while not lQRub.Eof do
        begin
          // Est-ce que le compte général appartient à la rubrique ?
          if ChercheCompteDansRub(lQRub.FindField('RB_COMPTE1').AsString,Compte) then
            // Est-ce que le compte général est exclu de la rubrique ?
            if not ChercheCompteDansRub(lQRub.FindField('RB_EXCLUSION1').AsString,Compte) then
            begin
              // Dans ce cas, le compte général appartient bien à la rubrique lQRub.FindField('RB_RUBRIQUE').AsString
              // On récupère donc la liste des sections associées à cette rubrique
              // FQ 21008 : ne pas prendre les sections fermées
              lstSQL := 'SELECT * FROM SECTION WHERE S_FERME="-" AND S_TRANCHEGENEDE="'+
                            lQRub.FindField('RB_RUBRIQUE').AsString+'"';
              // et on l'ajoute à la liste des section déjà trouvées
              lTSection.LoadDetailDBFromSQL('SECTION',lstSQL,True);
            end;
          lQRub.Next;
        end;
      finally
        Ferme (lQRub);
      end;

      for i := 0 to lTSection.Detail.Count - 1 do
        begin
        // gestion du changement d'axe
        if lStAxe <> lTSection.Detail[i].GetString('S_AXE') then
          begin
          if lTobVent.Detail.count > 0 then
            begin
            // Calcul des %
            _finaliseVentil ;
            // Enregistrement
            lTobVent.InsertDB( nil ) ;
            // Validation Axe
            result := True ;
            AxeVent[ lInAxe-1 ] := true;
            lTobVent.ClearDetail ;
            end ;
          // Reinit des variables
          lInNb     := 1 ;
          lStAxe    := lTSection.Detail[i].GetString('S_AXE') ;
          lInAxe    := ValeurI( lStAxe[2] ) ;
          lTotalUO  := 0 ;
          end ;

        // Axe valide ? (avec une section d'attente)
        SQL := 'SELECT ##TOP 1## 1 FROM VENTIL ' +
               'WHERE V_NATURE = "GE' + IntToStr( lInAxe ) + '" ' +
               'AND V_COMPTE = "' + Compte + '" ' +
               'AND V_NUMEROVENTIL = "' + IntToStr(lInNb) + '"';


        if SectAtt[ lInAxe-1 ] and not ExisteSQL(SQL) then
          begin
          lVentilUO := lTSection.Detail[i].GetDouble('S_UO') ;
          if lVentilUO = 0 then
            lVentilUO := 1.0 ;

          // Créatin de la ligne de ventil
          lTobL := TOB.Create('VENTIL', lTobVent, -1);

          lTobL.putValue('V_NATURE',        'GE' + IntToStr( lInAxe ) ) ;
          lTobL.putValue('V_COMPTE',        Compte ) ;
          lTobL.putValue('V_SECTION',       lTSection.Detail[i].GetString('S_SECTION')) ;
          lTobL.putValue('V_NUMEROVENTIL',  lInNb ) ;
          lTobL.putValue('V_MONTANT',       lVentilUO ) ;

          lTotalUO := lTotalUO + lVentilUO ;
          Inc( lInNb ) ;
          { FQ 21211 BVE 24.10.07 }
          end
          else
          begin
             AxeVent[ lInAxe-1 ] := true;
          { END FQ 21211 }
          end ;

        end ;
      lTSection.Free ;

      // gestion dernier Axe
      if lTobVent.Detail.count > 0 then
        begin
        // Calcul des %
        _finaliseVentil ;
        // Enregistrement
        lTobVent.InsertDB( nil ) ;
        result := True ;
        AxeVent[ lInAxe-1 ] := true;
        end ;

      CommitTrans ;

      Except
        on E : Exception do
          begin
          result := False ;
          RollBack;
          {$IFNDEF EAGLSERVER}
          PgiError( TraduireMemoire('Erreur à l''initialisation de la ventilation par défaut du compte.')
                    + ' (' + E.Message + ')') ;
          {$ENDIF}
          end ;

      end ;

    finally

      if Assigned( lTobVent )  then
        Freeandnil( lTobVent ) ;
      end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/10/2006
Modifié le ... : 15/10/2007
Description .. : FQ 21624 - CA - 15/10/2007 : contrôle de comptes exclus 
Suite ........ : dans les rubriques TRA
Mots clefs ... : 
*****************************************************************}
function ChercheCompteDansRub(CptRub,ComptGene : String) : Boolean ;
var
  St,
  StTemp : String ;
begin
  Result:=False ;
  while CptRub <> '' do begin
    St := ReadTokenSt(CptRub) ;
    StTemp := '' ;

    if St = '' then Continue ;
    if Pos('(', St) > 0 then St := Copy(St, 1, Pos('(', St)-1) ;

    if Pos(':',St)>0 then begin
      StTemp := Copy(St, Pos(':',St)+1, 200) ;
      System.Delete(St, Pos(':',St), 200) ;
    end ;

    if StTemp<>'' then begin
      St := BourrelaDonc(St, fbGene) ;
      StTemp := BourreLaDonc(StTemp, fbGene) ;
      if(ComptGene >= St) and (ComptGene <= StTemp) then begin
        Result := True ;
        Exit ;
      end ;
    end

    else begin
      // if Length(St) <= GetInfoCpta(fbaxe1).Lg then
      if Length(St) <= GetInfoCpta(fbGene).Lg then  // FQ 21624 - CA - 15/10/2007
        if(Pos(St,ComptGene) = 1) or (St = ComptGene) then begin
          Result := True ;
          Exit;
        end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

{JP 24/10/07 : Retourne les infos d'un axe en tenant compte du partage et du dossier
{---------------------------------------------------------------------------------------}
function GetInfosAxe(Axe : Integer; aDossier : string = '') : TInfoCpta;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  S : string;
  {$IFNDEF NOVH}
  F : TFichierBase;
  {$ENDIF NOVH}
begin
  {$IFNDEF NOVH}
  if EstTablePartagee('AXE') or not EstMultiSoc then begin
    F := AxeToFb('A' + IntToStr(Axe));
    Result.Lg           := VH^.Cpta[F].Lg;
    Result.Chantier     := VH^.Cpta[F].Chantier;
    Result.Structure    := VH^.Cpta[F].Structure;
    Result.Attente      := VH^.Cpta[F].Attente;
    Result.AxGenAttente := VH^.Cpta[F].AxGenAttente;
    Exit;
  end;
  {$ENDIF NOVH}

  Q := OpenSQL('SELECT * FROM ' + GetTableDossier(aDossier, 'AXE') + ' WHERE X_AXE = "' +
               'A' + IntToStr(Axe) + '"', True, -1, 'AXE' + IntToStr(Axe) + aDossier,true);
  if not Q.EOF then begin
    Result.Lg := Q.FindField('X_LONGSECTION').AsInteger;
    S := Q.FindField('X_BOURREANA').AsString;
    if S <> '' then Result.Cb := S[1]
               else Result.Cb := '0';
    Result.Chantier     := Q.FindField('X_CHANTIER').AsString = 'X';
    Result.Structure    := Q.FindField('X_STRUCTURE').AsString = 'X';
    Result.Attente      := Q.FindField('X_SECTIONATTENTE').AsString;
    Result.AxGenAttente := Q.FindField('X_GENEATTENTE').AsString;
  end;
  Ferme(Q);
end;

{------------------------------------------------------------------------------------
Ajout me
Déplacement des fonctions du AGIStdCompta, LetBatch,CPGeneraux_TOM dans ulibanalytique
en EAglServer
-------------------------------------------------------------------------------------}

function  CountNbEcriture ( LeCpte : String; Where : String; aDossier : string = '') : integer;
var
  SQL : String;
  Q   : TQuery;
begin
  Result := 0;
  SQL := 'SELECT COUNT(*) NBECRITURES ' +
         'FROM ' + GetTableDossier(aDossier, 'ECRITURE') + ' ' +
         'WHERE E_GENERAL = "' + LeCpte + '" ' +
         'AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' +
         Where ;
  Q := OpenSQL(SQL,true,-1,'',true);
  if not(Q.Eof) then
     Result := Q.FindField('NBECRITURES').AsInteger;
  Ferme(Q);
end;

Procedure EcrToAna ( OEcr,OAna : TOBM ) ;
BEGIN
OAna.PutMvt('Y_GENERAL',OEcr.GetMvt('E_GENERAL'))         ; OAna.PutMvt('Y_DATECOMPTABLE',OEcr.GetMvt('E_DATECOMPTABLE')) ;
OAna.PutMvt('Y_PERIODE',OEcr.GetMvt('E_PERIODE'))         ; OAna.PutMvt('Y_SEMAINE',OEcr.GetMvt('E_SEMAINE')) ;
OAna.PutMvt('Y_NUMEROPIECE',OEcr.GetMvt('E_NUMEROPIECE')) ; OAna.PutMvt('Y_NUMLIGNE',OEcr.GetMvt('E_NUMLIGNE')) ;
OAna.PutMvt('Y_EXERCICE',OEcr.GetMvt('E_EXERCICE'))       ; OAna.PutMvt('Y_NATUREPIECE',OEcr.GetMvt('E_NATUREPIECE')) ;
OAna.PutMvt('Y_QUALIFPIECE',OEcr.GetMvt('E_QUALIFPIECE')) ; OAna.PutMvt('Y_TYPEANALYTIQUE','-')                       ;
OAna.PutMvt('Y_VALIDE',OEcr.GetMvt('E_VALIDE'))           ; OAna.PutMvt('Y_ETABLISSEMENT',OEcr.GetMvt('E_ETABLISSEMENT')) ;
OAna.PutMvt('Y_ETAT',OEcr.GetMvt('E_ETAT'))               ; OAna.PutMvt('Y_DEVISE',OEcr.GetMvt('E_DEVISE')) ;
OAna.PutMvt('Y_TAUXDEV',OEcr.GetMvt('E_TAUXDEV'))         ; OAna.PutMvt('Y_TYPEANOUVEAU',OEcr.GetMvt('E_TYPEANOUVEAU')) ;
OAna.PutMvt('Y_DATETAUXDEV',OEcr.GetMvt('E_DATETAUXDEV')) ; OAna.PutMvt('Y_ECRANOUVEAU',OEcr.GetMvt('E_ECRANOUVEAU')) ;
OAna.PutMvt('Y_REFINTERNE',OEcr.GetMvt('E_REFINTERNE'))   ; OAna.PutMvt('Y_LIBELLE',OEcr.GetMvt('E_LIBELLE')) ;
OAna.PutMvt('Y_JOURNAL',OEcr.GetMvt('E_JOURNAL'))         ;
// V9 CEGID
OAna.PutMvt('Y_DATPER',iDate1900) ;
OAna.PutMvt('Y_ENTITY',0) ;
OAna.PutMvt('Y_REFGUID','') ;
END ;

Procedure FiniMontantsAna ( OAna : TOBM ; RY : R_YV ; DEV : RDEVISE ; UnSeul : boolean; aDossier : string = '') ;
Var DD,CD,DP,CP : Double ;
BEGIN
OAna.PutMvt('Y_TOTALDEVISE',RY.TotDevise) ; OAna.PutMvt('Y_TOTALECRITURE',RY.TotPivot) ;
OAna.PutMvt('Y_SECTION',RY.Section)        ; OAna.PutMvt('Y_NUMVENTIL',RY.NumVentil) ;
OAna.PutMvt('Y_POURCENTAGE',RY.Pourcentage) ;
//SG6 03.03.05 Gestion croisaxe
if GetParamSocDossierMS('SO_CROISAXE', False, aDossier) then
begin
  OAna.PutMvt('Y_SOUSPLAN1',RY.SousPlan1);
  OAna.PutMvt('Y_SOUSPLAN2',RY.SousPlan2);
  OAna.PutMvt('Y_SOUSPLAN3',RY.SousPlan3);
  OAna.PutMvt('Y_SOUSPLAN4',RY.SousPlan4);
  OAna.PutMvt('Y_SOUSPLAN5',RY.SousPlan5);
end;

if OAna.GetMvt('Y_NUMVENTIL')=1 then OAna.PutMvt('Y_TYPEMVT','AE') else  OAna.PutMvt('Y_TYPEMVT','AL') ;
if UnSeul then
   BEGIN
   if RY.Debit then
      BEGIN
      DD:=RY.TotDevise ; DP:=RY.TotPivot ; CP:=0 ; CD:=0 ;
      END else
      BEGIN
      CD:=RY.TotDevise ; CP:=RY.TotPivot ; DP:=0 ; DD:=0 ;
      END ;
   OAna.PutMvt('Y_DEBITDEV',DD)  ; OAna.PutMvt('Y_CREDITDEV',CD) ;
   OAna.PutMvt('Y_DEBIT',DP)     ; OAna.PutMvt('Y_CREDIT',CP) ;
   END else
   BEGIN
   if RY.Debit then
      BEGIN
      DD:=Arrondi(RY.TotDevise*RY.Pourcentage/100.0,DEV.Decimale) ; CD:=0 ;
      END else
      BEGIN
      CD:=Arrondi(RY.TotDevise*RY.Pourcentage/100.0,DEV.Decimale) ; DD:=0 ;
      END ;
   OAna.PutMvt('Y_DEBITDEV',DD)  ; OAna.PutMvt('Y_CREDITDEV',CD) ;
   if DEV.Code<>V_PGI.DevisePivot then
      BEGIN
      DP:=DeviseToPivot(DD,Dev.Taux,Dev.Quotite) ; CP:=DeviseToPivot(CD,Dev.Taux,Dev.Quotite) ;
      OAna.PutMvt('Y_DEBIT',DP) ; OAna.PutMvt('Y_CREDIT',CP) ;
      END else
      BEGIN
      OAna.PutMvt('Y_DEBIT',DD) ; OAna.PutMvt('Y_CREDIT',CD) ;
      END ;
   END ;
END ;

Procedure EquilibreDerniere ( TVentAna : TList ; OEcr : TOBM ; RY : R_YV ; DEV : RDevise ; Force : boolean; aDossier : string = '') ;
Var SommeDevise,SommePivot,SommePourc : Double ;
    i : integer ;
    OAna : TOBM ;
    Debit : Boolean ;
    DD,CD,DP,CP : Double ;
BEGIN
SommeDevise:=0 ; SommePivot:=0 ; SommePourc:=0 ;
for i:=0 to TVentAna.Count-1 do
    BEGIN
    OAna:=TOBM(TVentAna[i]) ;
    SommeDevise:=SommeDevise+OAna.GetMvt('Y_DEBITDEV')+OAna.GetMvt('Y_CREDITDEV') ;
    SommePivot:=SommePivot+OAna.GetMvt('Y_DEBIT')+OAna.GetMvt('Y_CREDIT') ;
    SommePourc:=SommePourc+OAna.GetMvt('Y_POURCENTAGE') ;
    END ;
if ((Force) or (Arrondi(SommePourc-100.0,ADecimP)=0)) then
   BEGIN
   {Equilibre sur la dernière}
   OAna:=TOBM(TVentAna[TVentAna.Count-1]) ;
   DD:=OAna.GetMvt('Y_DEBITDEV') ; CD:=OAna.GetMvt('Y_CREDITDEV') ;
   DP:=OAna.GetMvt('Y_DEBIT') ; CP:=OAna.GetMvt('Y_CREDIT') ;
   Debit:=(DD<>0) ;
   if Debit then
      BEGIN
      OAna.PutMvt('Y_DEBIT',DP+RY.TotPivot-SommePivot) ;
      OAna.PutMvt('Y_DEBITDEV',DD+RY.TotDevise-SommeDevise) ;
      OAna.PutMvt('Y_CREDIT',0) ; OAna.PutMvt('Y_CREDITDEV',0) ;
      END else
      BEGIN
      OAna.PutMvt('Y_CREDIT',CP+RY.TotPivot-SommePivot) ;
      OAna.PutMvt('Y_CREDITDEV',CD+RY.TotDevise-SommeDevise) ;
      OAna.PutMvt('Y_DEBIT',0) ; OAna.PutMvt('Y_DEBITDEV',0) ;
      END ;
   END else
   BEGIN
   {Equilibre sur une nouvelle ligne sur section attente}
   OAna:=TOBM.Create(EcrAna,RY.Axe,True) ; EcrToAna(OEcr,OAna) ;
   RY.Section:=GetInfosAxe(ValeurI(Copy(RY.Axe, 2, 1)), aDossier).Attente ; {JP 24/10/07 : Multi sociétés et partage}
   RY.Pourcentage:=Arrondi(100.0-SommePourc,ADecimP) ;
   RY.NumVentil:=TVentAna.Count+1 ;
   OAna.PutMvt('Y_SECTION',RY.Section) ; OAna.PutMvt('Y_POURCENTAGE',RY.Pourcentage) ;
   OAna.PutMvt('Y_NUMVENTIL',RY.NumVentil) ;
   FiniMontantsAna(OAna,RY,DEV,False) ;
   TVentAna.Add(OAna) ;
   Equilibrederniere(TVentAna,OEcr,RY,DEV,True, aDossier) ;
   END ;
END ;           

Procedure VentileGenerale ( Ax : String3 ; OEcr : TOBM ; DEV : RDevise ; TVentAna : TList ; Preventil : boolean; aDossier : string = '' ) ;
Var OAna     : TOBM ;
    NumAxe   : integer ;
    RY       : R_YV ;
    Debit    : boolean ;
    Q        : TQuery ;
    CpteGene,StV : String ;
    NatP         : String3 ;
    TotPivot,TotDevise : Double ;
    InfoAxe : TInfoCpta;
BEGIN
NumAxe:=StrToInt(Copy(Ax,2,1)) ; CpteGene:=OEcr.GetMvt('E_GENERAL') ;
TotPivot:=OEcr.GetMvt('E_DEBIT')+OEcr.GetMvt('E_CREDIT') ; Debit:=(OEcr.GetMvt('E_DEBIT')<>0) ;
NatP:=OEcr.GetMvt('E_NATUREPIECE') ;
if NatP<>'ECC' then TotDevise:=OEcr.GetMvt('E_DEBITDEV')+OEcr.GetMvt('E_CREDITDEV')
               else TotDevise:=TotPivot ;
StV:='SELECT * FROM ' + GetTableDossier(aDossier, 'VENTIL') + ' WHERE V_NATURE="GE'+IntToStr(NumAxe)+'" AND V_COMPTE="'+CpteGene+'" '
    +'ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
Q:=OpenSQL(StV,True,-1,'',true) ;
if ((Not Q.EOF) and (Preventil)) then
   BEGIN
   While Not Q.EOF do
      BEGIN
      if GetParamSocDossierMS('SO_CROISAXE', False, aDossier) then
      begin
        RY.SousPlan1 := Q.FindField('V_SOUSPLAN1').AsString;
        RY.SousPlan2 := Q.FindField('V_SOUSPLAN2').AsString;
        RY.SousPlan3 := Q.FindField('V_SOUSPLAN3').AsString;
        RY.SousPlan4 := Q.FindField('V_SOUSPLAN4').AsString;
        RY.SousPlan5 := Q.FindField('V_SOUSPLAN5').AsString;
      end
      else
      begin
      end;
      RY.Axe:=Ax ; RY.Section:=Q.FindField('V_SECTION').AsString ;
      RY.Pourcentage:=Q.FindField('V_TAUXMONTANT').AsFloat ;
      RY.NumVentil:=Q.FindField('V_NUMEROVENTIL').AsInteger ;
      RY.TotDevise:=TotDevise ; RY.TotPivot:=TotPivot ; RY.Debit:=Debit ;
      if RY.Section<>'' then
         BEGIN
         OAna:=TOBM.Create(EcrAna,Ax,True) ;
         EcrToAna(OEcr,OAna) ; FiniMontantsAna(OAna,RY,DEV,False, aDossier) ;
         TVentAna.Add(OAna) ;
         END ;
      Q.Next ;
      END ;
   EquilibreDerniere(TVentAna,OEcr,RY,DEV,False, aDossier) ;
   END else
   BEGIN
   {JP 24/10/07 : FQ 21714 : gestion du multi sociétés et du partage du référentiel}
   InfoAxe := GetInfosAxe(NumAxe, aDossier);
   RY.Axe:=Ax ;
   RY.Section := InfoAxe.Attente ;
   RY.NumVentil:=1 ; RY.Pourcentage:=100.0 ;
   RY.TotDevise:=TotDevise ; RY.TotPivot:=TotPivot ; RY.Debit:=Debit ;
   if GetParamSocDossierMS('SO_CROISAXE', False, aDossier) then
   begin
     if GetParamSocDossierMS('SO_VENTILA1', False, aDossier) then begin
       InfoAxe := GetInfosAxe(1, aDossier);
       RY.SousPlan1 := InfoAxe.Attente;
     end
     else
       RY.SousPlan1 := '';

     if GetParamSocDossierMS('SO_VENTILA2', False, aDossier) then begin
       InfoAxe := GetInfosAxe(2, aDossier);
       RY.SousPlan2 := InfoAxe.Attente;
     end
     else
       RY.SousPlan2 := '';

     if GetParamSocDossierMS('SO_VENTILA3', False, aDossier) then begin
       InfoAxe := GetInfosAxe(3, aDossier);
       RY.SousPlan3 := InfoAxe.Attente;
     end
     else
       RY.SousPlan3 := '';

     if GetParamSocDossierMS('SO_VENTILA4', False, aDossier) then begin
       InfoAxe := GetInfosAxe(4, aDossier);
       RY.SousPlan4 := InfoAxe.Attente;
     end
     else
       RY.SousPlan4 := '';

     if GetParamSocDossierMS('SO_VENTILA5', False, aDossier) then begin
       InfoAxe := GetInfosAxe(5, aDossier);
       RY.SousPlan5 := InfoAxe.Attente;
     end
     else
       RY.SousPlan5 := '';
   end;
   if RY.Section<>'' then
      BEGIN
      OAna:=TOBM.Create(EcrAna,Ax,True) ;
      EcrToAna(OEcr,OAna) ; FiniMontantsAna(OAna,RY,DEV,True, aDossier) ;
      TVentAna.Add(OAna) ;
      END ;
   END ;
Ferme(Q) ;
END ;

function VentileCompte( LeCpte : String ; Where : String ; VentilAxe : array of string ; PreVentil : boolean; aDossier : string = '') : boolean;
var
  Dev       : RDevise ;
  Q         : TQuery ;
  QAna      : TOB ;
  OEcr      : TOBM ;
  OAna      : TOBM ;
  TVentAna  : TList ;
  NewAxe    : String ;
  i,k       : Integer ;
  Okok      : Boolean ;
  SQL       : String;
begin
  Okok := False ;
  SQL := 'SELECT * ' +
         'FROM ' + GetTableDossier(aDossier, 'ECRITURE') + ' ' +
         'WHERE E_GENERAL = "' + LeCpte + '" ' +
         'AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")' +
         Where ;
  Q := OpenSQL(SQL,true,-1,'',true);
  try
    QAna := TOB.Create('ANALYTIQ',nil,-1);
    OEcr     := TOBM.Create(EcrGen,'',False) ;
    TVentAna := TList.Create ;
    try
      While not Q.EOF do
      begin
        OEcr.ChargeMvt(Q) ;
        Dev.Code     := Q.FindField('E_DEVISE').AsString ;
        GetInfosDevise(Dev) ;
        Dev.DateTaux := Q.FindField('E_DATETAUXDEV').AsDateTime ;
        Dev.Taux     := Q.FindField('E_TAUXDEV').AsFloat ;

        //SG6 30/12/2004 Gestion ventilation classique
        if not GetParamSocDossierMS('SO_CROISAXE', False, aDossier) then
        begin
           for i:=1 to MaxAxe do
           if  (VentilAxe[i] = 'X') then
           begin
              NewAxe := 'A' + IntToStr(i) ;
              SQL := 'SELECT Y_GENERAL ' +
                     'FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' ' +
                     'WHERE Y_GENERAL = "' + LeCpte + '" AND Y_AXE = "' + NewAxe + '" ' +
                     'AND Y_JOURNAL = "' + Q.FindField('E_JOURNAL').AsString + '" ' +
                     'AND Y_EXERCICE = "' + Q.FindField('E_EXERCICE').AsString + '" ' +
                     'AND Y_DATECOMPTABLE = "' + UsDateTime(Q.FindField ('E_DATECOMPTABLE').asDatetime) + '" ' +
                     'AND Y_QUALIFPIECE = "' + Q.FindField('E_QUALIFPIECE').AsString + '" ' +
                     'AND Y_NUMEROPIECE = ' + Q.FindField ('E_NUMEROPIECE').AsString + ' ' +
                     'AND Y_NUMLIGNE = '+ Q.FindField ('E_NUMLIGNE').AsString ;
              if (Q.FindField ('E_ANA').asstring = 'X') and ExisteSQl(SQL) then continue;
              VideListe(TVentAna) ;
              Okok := True ;
              VentileGenerale( NewAxe, OEcr, Dev, TVentAna, True, aDossier) ;
              for k:=0 to TVentAna.Count-1 do
              begin
                 OAna := TOBM(TVentAna[k]) ;
                 OAna.EgalChampsTOB(QAna);
                 {JP 23/10/07 : FQ 21714 : gestion du multi sociétés
                  QAna.InsertDB(nil,False);}
                 InsertTobMS(QAna, aDossier);
                 MajSoldeSectionTOB(QAna,True, aDossier);
              end;
           end;
        end
        //SG6 30/12/2004 Gestion ventilations croisaxe
        else
        begin
           NewAxe:='A'+IntToStr(RecherchePremDerAxeVentil.premier_axe);
           SQL := 'SELECT Y_GENERAL ' +
                  'FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' ' +
                  'WHERE Y_GENERAL = "' + LeCpte + '" AND Y_AXE = "' + NewAxe + '" ' +
                  'AND Y_JOURNAL = "' + Q.FindField('E_JOURNAL').AsString + '" ' +
                  'AND Y_EXERCICE = "' + Q.FindField('E_EXERCICE').AsString + '" ' +
                  'AND Y_DATECOMPTABLE = "' + UsDateTime(Q.FindField ('E_DATECOMPTABLE').asDatetime) + '" ' +
                  'AND Y_QUALIFPIECE = "' + Q.FindField('E_QUALIFPIECE').AsString + '" ' +
                  'AND Y_NUMEROPIECE = ' + Q.FindField ('E_NUMEROPIECE').AsString + ' ' +
                  'AND Y_NUMLIGNE = '+ Q.FindField ('E_NUMLIGNE').AsString ;
           if (Q.FindField ('E_ANA').AsString = 'X') and ExisteSQl(SQL) then continue;
           VideListe(TVentAna) ;
           Okok   := True ;
           VentileGenerale(NewAxe, OEcr, Dev, TVentAna, PreVentil, aDossier) ;
           for k:=0 to TVentAna.Count-1 do
           begin
              OAna := TOBM(TVentAna[k]) ;
              OAna.EgalChampsTOB(QAna);
              if not(PreVentil) then
              begin
                 for i:=1 to MaxAxe do
                 begin
                    if (VentilAxe[i] = 'X') then
                    begin
                       // On tiens pas compte d'eventuelles ventilation type //TEMPORAIRE
                       OAna.PutMvt('Y_SOUSPLAN',GetInfoCpta (AxeToFb('A'+IntToStr(i))).Attente);
                    end;
                 end;
              end;
             {JP 23/10/07 : FQ 21714 : gestion du multi sociétés
              QAna.InsertDB(nil,False);}
             InsertTobMS(QAna, aDossier);
             MajSoldeSectionTOB(QAna,True, aDossier);
           end;
        end;
        Q.Next ;
      end;
      if Okok then
        ExecuteSQL('UPDATE ' + GetTableDossier(aDossier, 'ECRITURE') + ' SET E_ANA = "X" ' +
                   'WHERE E_GENERAL = "' + LeCpte + '" ' +
                   'AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") ' +
                   Where);
    finally
      Ferme(Q);
      if Assigned(QAna) then FreeAndNil(QAna);
      VideListe(TVentAna);
      if Assigned(TVentAna) then FreeAndNil(TVentAna);
      if Assigned(OEcr) then FreeAndNil(OEcr);
    end;
  except
     on E: Exception do
     begin
{$IFNDEF EAGLSERVER}
        PgiError('Erreur durant VentileCompte : ' + E.Message );
{$ENDIF}
     end;
  end; 
  Result := Okok;
end;


procedure MajVentilation(LeCpte: String ; VentilAxe : array of string ; PreVentil : Boolean; aDossier : string = '');
var
  SQL       : String;
  Where     : String;
  QExercice : TQuery;
  QJournal  : TQuery;
  Exercice  : String;
  Journal   : String;
  Periode   : integer;
  PeriodeFin: integer;
  Ok        : boolean;
  i,j       : Integer;
  vTobAna   : TOB;
  NewAxe    : string;
  NbEcriture: integer;
begin
  { BVE 04.05.07 On segmente la requete par Exercice, Periode et Journal
    pour eviter les problèmes sur les grosses bases de données    }
  Ok := false;
  try
    BeginTrans;
    NbEcriture := CountNbEcriture(LeCpte, '', aDossier);
    if NbEcriture > MaxEcriture then
    begin
      SQL := 'SELECT EX_EXERCICE, EX_DATEDEBUT, EX_DATEFIN ' +
             'FROM ' + GetTableDossier(aDossier, 'EXERCICE') + ' ' + {JP 23/10/07 : FQ 21714}
             'ORDER BY EX_DATEDEBUT, EX_DATEFIN';
             
      QExercice := OpenSQL(SQL,true,-1,'',true);
      while not(QExercice.EOF) do
      begin
        // Pour chaque Exercice
        Exercice := QExercice.FindField('EX_EXERCICE').AsString;
        Where := ' AND E_EXERCICE = "' + Exercice + '"';
        NbEcriture := CountNbEcriture(LeCpte, Where, aDossier); {JP 23/10/07 : FQ 21714}
        if NbEcriture > MaxEcriture then
        begin
          // On recupere toutes les périodes de l'exercice courant
          Periode      := GetPeriode(QExercice.FindField('EX_DATEDEBUT').AsDateTime);
          PeriodeFin   := GetPeriode(QExercice.FindField('EX_DATEFIN'  ).AsDateTime);
          while (Periode <= PeriodeFin) do
          begin
             // Pour chaque Période de l'exercice
             Where := ' AND E_EXERCICE = "' + Exercice + '"' +
                      ' AND E_PERIODE = "' + IntToStr(Periode) + '"';
             NbEcriture := CountNbEcriture(LeCpte, Where, aDossier); {JP 23/10/07 : FQ 21714}
             if NbEcriture > MaxEcriture then
             begin
               // On recupere les journaux ouvert durant la période
               SQL := 'SELECT J_JOURNAL ' +
                      'FROM ' + GetTableDossier(aDossier, 'JOURNAL') + ' ' + {JP 23/10/07 : FQ 21714}
                      'WHERE (J_DATEOUVERTURE < "' + usDateTime(GetDateTimeFromPeriode(IncPeriode(Periode))) + '" )' +
                      'AND ((J_DATEFERMETURE >= "' + usDateTime(GetDateTimeFromPeriode(Periode)) + '" )' +
                      'OR (J_DATEFERMETURE = "' + usDateTime(iDate1900) + '" ))';
               QJournal := OpenSQL(SQL,true,-1,'',true);
               while not(QJournal.EOF) do
               begin
                  // Pour chaque journal
                  Journal := QJournal.FindField('J_JOURNAL').AsString;
                  Where := ' AND E_EXERCICE = "' + Exercice + '"' +
                           ' AND E_PERIODE = "' + IntToStr(Periode) + '"' +
                           ' AND E_JOURNAL = "' + Journal + '"';
                  // On execute le traitement
                  // TRAITEMENT
                  Ok := Ok or VentileCompte(LeCpte,Where,VentilAxe,PreVentil, aDossier); {JP 23/10/07 : FQ 21714}
                  // On passe au journal suivant
                  QJournal.Next;
               end;
               Ferme(QJournal);
             end
             else
             begin
                // TRAITEMENT
                if NbEcriture > 0 then
                   Ok := Ok or VentileCompte(LeCpte,Where,VentilAxe,PreVentil, aDossier); {JP 23/10/07 : FQ 21714}
             end;
             // On passe à la période suivante
             Periode := IncPeriode(Periode);
          end;
        end
        else
        begin
           // TRAITEMENT
           if NbEcriture > 0 then
              Ok := Ok or VentileCompte(LeCpte,Where,VentilAxe,PreVentil, aDossier); {JP 23/10/07 : FQ 21714}
        end;
        // On passe à l'exercice suivant
        QExercice.Next
      end;
      Ferme(QExercice);
    end
    else
    begin
       // TRAITEMENT
       if NbEcriture > 0 then
          Ok := Ok or VentileCompte(LeCpte,'',VentilAxe,PreVentil, aDossier); {JP 23/10/07 : FQ 21714}
    end;

    if not Ok then
    begin
       // pas de ventilation
       if not GetParamSocDossierMS('SO_CROISAXE', False, aDossier) then {JP 23/10/07 : FQ 21714}
       begin
          for i:=1 to MaxAxe do
          begin
             if not (VentilAxe[i] = 'X')  then
             begin
                NewAxe := 'A'+IntToStr(i) ;
                //SG6 03.03.05 Maj Solde
                vTobAna := TOB.Create('$ANA', nil, -1);
                SQL := 'SELECT * ' +
                       'FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' ' + {JP 23/10/07 : FQ 21714}
                       'WHERE Y_GENERAL = "' + LeCpte + '" ' +
                       'AND Y_AXE = "' + NewAxe + '" ';
                vTobAna.LoadDetailDBFromSQL('ANALYTIQUE', SQL);
                for j := 1 to vTobAna.Detail.Count - 1 do MajSoldeSectionTOB(vTobAna.Detail[j], False, aDossier);
                vTobAna.Free;
                {JP 23/10/07 : FQ 21714}
                ExecuteSQL ('DELETE FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' WHERE Y_GENERAL="'+LeCpte+'" and Y_AXE="'+NewAxe+'"');
                ExecuteSQL ('DELETE FROM ' + GetTableDossier(aDossier, 'VENTIL') + ' WHERE V_COMPTE="'+LeCpte+'" And V_NATURE="GE'+IntToStr(i)+'"');
                ExecuteSQL ('DELETE FROM ' + GetTableDossier(aDossier, 'VENTIL') + ' WHERE V_COMPTE="'+LeCpte+'" And V_NATURE="IM'+IntToStr(i)+'"');
             end;
             if not (VentilAxe[0] = 'X') then
             begin
                ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE="-"'+' WHERE G_GENERAL="'+ LeCpte+'"') ;
                ExecuteSQL('UPDATE ' + GetTableDossier(aDossier, 'ECRITURE') + ' SET E_ANA="-" WHERE E_GENERAL="'+LeCpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")') ;
                //SG6 03.03.05 Maj Solde
                vTobAna:=TOB.Create('$ANA', nil, -1);
                SQL := 'SELECT * ' +
                       'FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' ' + {JP 23/10/07 : FQ 21714}
                       'WHERE Y_GENERAL = "' + LeCpte + '" ';
                vTobAna.LoadDetailDBFromSQL('ANALYTIQUE', SQL);
                for j := 1 to vTobAna.Detail.Count - 1 do MajSoldeSectionTOB(vTobAna.Detail[j], False, aDossier);
                vTobAna.Free;
                ExecuteSQL('DELETE FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' WHERE Y_GENERAL = "' + LeCpte + '"');
             end;
          end;
       end
       else
       begin
          // YMO 10/02/2006 FQ16434 LE compte était toujours mis en non-ventilable
          if not (VentilAxe[0] = 'X') then
          begin
             ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE="-"'+' where G_GENERAL="'+ LeCpte+'"') ;
             ExecuteSQL('UPDATE ' + GetTableDossier(aDossier, 'ECRITURE') + ' SET E_ANA="-" WHERE E_GENERAL="'+LeCpte+'" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")') ;
             //SG6 03.03.05 Maj Solde
             vTobAna := TOB.Create('$ANA', nil, -1);
             vTobAna.LoadDetailDBFromSQL('ANALYTIQUE', 'SELECT * FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' WHERE Y_GENERAL = "' + LeCpte + '"');
             for j := 1 to vTobAna.Detail.Count - 1 do MajSoldeSectionTOB(vTobAna.Detail[j], False, aDossier);
             vTobAna.Free;
             ExecuteSQL ('DELETE FROM ' + GetTableDossier(aDossier, 'ANALYTIQ') + ' WHERE Y_GENERAL = "' + LeCpte + '"');
          end;
       end;
    end;
    // Tout c'est bien déroulé on valide la transaction
    CommitTrans
  except
    // Il y a eu un soucis
    on E: Exception do
    begin
      // On remet la BDD en état
      RollBack;
      if assigned(QExercice) then Ferme(QExercice);
      if assigned(QJournal) then Ferme(QJournal);
{$IFNDEF EAGLSERVER}
      PgiError('Erreur durant MajVentilation : ' + E.Message );
{$ENDIF}
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// VLA 26/01/2005
procedure MajGVentil(OKgene : Boolean; LeCpte : string);
var
    Where   : string;
    Tableau : array[0..MaxAxe] of string;
    i       : integer;
    Q1      : TQuery;
    SectAtt : Array[1..5] of Boolean ; // stockage des section d'attente
    lAxeVent: Array[1..5] of Boolean ; // résultat de la ventilation auto // FQ 21259 SBO 27/08/2007
begin
  if OkGene then
  begin
     Where := 'WHERE (G_VENTILABLE1="X" OR G_VENTILABLE2="X" '+
              'OR G_VENTILABLE3="X" OR G_VENTILABLE4="X" OR G_VENTILABLE5="X")';
     if ExisteSQL ('SELECT * From GENERAUX '+ Where) then
        ExecuteSQL('UPDATE  GENERAUX SET G_VENTILABLE="X" '+Where) ;
     end
     else
     begin
     {   for i := 0 to MaxAxe do
        begin
           if AnaEstVentilable(i,LeCpte) then
              Tableau[i] := 'X'
           else
              Tableau[i] := '-';
        end;   }
        Q1 := Opensql ('SELECT G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5 From GENERAUX Where G_GENERAL="'+LeCpte+'"', TRUE,-1,'',true);
        Tableau[1] := Q1.FindField('G_VENTILABLE1').asString ;
        Tableau[2] := Q1.FindField('G_VENTILABLE2').asString ;
        Tableau[3] := Q1.FindField('G_VENTILABLE3').asString ;
        Tableau[4] := Q1.FindField('G_VENTILABLE4').asString ;
        Tableau[5] := Q1.FindField('G_VENTILABLE5').asString ;
        if (Q1.FindField('G_VENTILABLE1').asstring='X') or
           (Q1.FindField('G_VENTILABLE2').asstring='X') or
           (Q1.FindField('G_VENTILABLE3').asstring='X') or
           (Q1.FindField('G_VENTILABLE4').asstring='X') or
           (Q1.FindField('G_VENTILABLE5').asstring='X') then
            Tableau[0] := 'X'
             else
              Tableau[0] := '-' ;
        Ferme(Q1) ;

        // Détermination auto des paramètres de ventilations du comptes en fonction des paramètres des sections :
        // - fourchette de comptes
        // - Unités d'oeuvre
        for i := 1 to 5 do
           SectAtt[i] := (Tableau[i] = 'X');
        InitialiseVentilDefaut(LeCpte,SectAtt,lAxeVent); // FQ 21259 SBO 27/08/2007

        MajVentilation(LeCpte,Tableau,True);
     end;
end;


end.





