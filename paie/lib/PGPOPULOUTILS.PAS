{***********UNITE*************************************************
Auteur  ...... : na
Créé le ...... : 20/06/2006
Modifié le ... :   /  /
Description .. : PGPOPULOUTILS
                 - Création de la TOB salarié pour une population
                 - Recherche de la population à laquelle appartient un salarié
                 - Création de la TOB salarié pour un type de population
Mots clefs ... :
*****************************************************************
PT1  10/08/2007  FLO      Recalcul automatique des compteurs lors d'une modification de population
PT2  24/08/2007  FLO      Chargement dynamique des types de population
PT3  29/11/2007  NA       Historisation = Oui si date simul < date début période paie active (et non pas date fin période)
PT4  : 23/01/2008 GGU V_81 Lors de l'affichage des banques du salarié en vision SAV, on a des messages
PT6  10/07/2008  NA  V_850 FQ 15622 Si IDR non sérialisé , ne pas afficher de message de validation de la population
                           et ne pas contrôler la saisie des critères (pour les pop IDR)
}
Unit PGPOPULOUTILS;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Utob,
     PgOutils2;

//             LA DIRECTIVE DE COMPILATION {IDEMPOP} PERMET D'ACTIVER OU NON
//             LA GESTION DES CHAMPS Idem Population DANS LA FICHE SALARIE
//
  //###############   FONCTIONS ET OBJETS INTERNES   #####################
  //Ces fonctions ne doivent normalement pas être utilisées en dehors de la gestion des populations.

  //FONCTION INTERNE
  //Renvoie la liste des salariés d'une population (pour un type donnée)
  function salariepop(Codepop, population : string ; datesimul : Tdatetime) : TOB;
  //FONCTION INTERNE
  function pgsalarietypepop(typepop, predefini: string; daterech : Tdatetime) : TOB;
  //FONCTION INTERNE
  function pgpopsalarie(salarie, typepop : string) : string;
  //FONCTION OBSOLETE
  function pgdoublonpopsal(salarie, typepop : string): boolean;
  //Récupère la valeur d'un paramètre
  //A partir du codeparametre
//  function GetParametre(CodeParametre, Module, Champ : String; TypeParametre : String = 'POP') : String;
  //FONCTION INTERNE
  //Effectue la validation d'un paramétrage de population
  Function ValidPopulParametres(TypePop : String = '') : Boolean;
  //FONCTION INTERNE
  //Vérifie la validité d'un paramétrage de population dans les parametres sociétés
  Function CheckValidPopulParametres(TypePop : String = '') : Boolean;
  //FONCTION INTERNE
  //Renvoie la liste des champs impliqués dans le paramétrage des populations
  Function GetFieldsListFromPopulParam(TypePop : String = '') : TStringList;
  //FONCTION INTERNE
  //Renvoie vrai si l'un des champs passé en tableau fait parti des champs
  //impliqués dans le paramétrage des populations
  Function isIncludeInPopulParam(ListChamps : Array of String; TypePop : String = '') : Boolean;
  //FONCTION INTERNE
  //Renvoie les types de populations
  Function GetTypesPop : TStringList;  //PT2


  //Objet INTERNE
  //Le compteur de référence tiens à jour le nombre de références faites
  //pour un couple (Critère,Valeur de critère)
  Type
    TCompteurReference = class (TObject)
  public
    FCriteres : Array of integer;
    FValeursCriteres : Array of String;
    constructor Create;
    destructor Destroy; override;
    Procedure AddCritereValeur(IDENT : integer; valeur : String);
    Procedure AddCritere(IDENT : integer);
    Procedure Increase;
    Function GetValCritere(IDENT : integer) : String;
    Procedure SetValCritere(IDENT : integer; valeur : String);
    Function GetNbrReferences : Integer;
    Function CritereExiste(IDENT: integer) : Boolean;
  private
    FTypePopulation : String;
    FPredefini : String;
    FCompteur : Integer;
  end;

  //Objet INTERNE
  //La matrice des compteurs de référence tiens à jour la liste des critères
  //utilisés et de leur différentes valeurs possibles (et du nombre de référence
  //de chacun des couples (Critère,Valeur de critère) par l'intermédiaire du compteur)
  Type
    TMatriceCompteurReference = class (TObject)
  public
    TypePopulation : String;
    Predefini : String;
    constructor Create;
    destructor Destroy; override;
    Procedure AddCritere(IDENT : integer);
    Procedure AddCriteres(CriteresIDENT : Array of integer); Overload;
    Function CritereExiste(IDENT: integer) : Boolean;
    Function IsValide : Boolean;
    Function IsEmpty : Boolean;
    Function GetCompteur(ListeCriteres : Array of integer; ListeValeurs : Array of String) : TCompteurReference;
    Procedure UpdateCompteurs(ListeCriteres : Array of integer; ListeValeurs : Array of String);
  private
    FCriteres : Array of integer;
    ListeCompteurs : Array of TCompteurReference;
    Procedure AddCompteur(CriteresIDENT : Array of integer; valeurs : Array of String);
  end;

  //Objet INTERNE
  //Le Validator permet de valider un paramétrage en contrôlant l'unicité et
  //l'exhaustivité de celui-ci
  //Pour cela, on vérifie que chaque valeur possible de chaque critère utilisé est
  //référencée une fois et une seule dans le paramétrage
  Type
    TParamPopulValidator = class (TObject)
  public
    constructor Create;
    destructor Destroy; override;
    Function GetMatrice(TypePop, Predefini : String) : TMatriceCompteurReference;
    Procedure UpdateCompteurs;
    Function IsValide(TypePopulation : String) : Boolean;
  private
    ListeMatrice : Array of TMatriceCompteurReference;
    ListeTypes : TStringList; //PT2
    Procedure NewMatrice(TypePop, Predefini : String);
  end;


  //###############   FONCTIONS ET OBJETS PUBLIQUES   #####################
  //Fonction a utiliser pour autoriser ou non l'usage des populations
  //Cette fonction permet aussi de valider un type de population qui à été mis à jour
  Function CanUsePopulation(TypePop : String = '') : Boolean;

  //Fonction a utiliser pour "dévalider" un type de population en cas de changements
  //Met le paramsoc de validité d'une population à vrai ou faux
  Procedure SetValidePopulParamSoc(TypePop : String; Valide : Boolean);

  //Fonction qui renvoie le prédefini pris en compte par le paramétrage des populations
  //pour un type de populations
  Function GetPredefiniPopulation(TypePop : String) : String;

  //Objet de gestion des populations
  Type
    TUpdateIdemPop = class (TObject)
  public
    constructor Create;
    destructor Destroy; override;
    //Fonction d'initialisation/Mise à jour de la table d'association salarié-population
    function MajSALARIEPOPUL(typepop : string; daterech : Tdatetime; ListUpdatesChamps : Array of String) : TStringList;   Overload;
    function MajSALARIEPOPUL(typepop : string; daterech : Tdatetime) : TStringList;   Overload;
    //Fonction de mis à jour d'un salarié (populations associées et champs Idem Population
    //de la fiche salarié
    function MajSALARIEPOPULSalarie(Salarie: String; daterech : Tdatetime; ListUpdatesChamps : Array of String) : TStringList;   Overload;
    function MajSALARIEPOPULSalarie(Salarie: String; daterech : Tdatetime) : TStringList;   Overload;
    //Fonction de mis à jour d'une liste de salariés (uniquement les champs Idem Population de la fiche salarié )
    function MajDonneesSalaries(TobListSalarie: TOB; daterech : Tdatetime; TypePop : String = '') : TStringList;
    //Fonction de mis à jour des salariés d'une population (uniquement les champs Idem Population de la fiche salarié )
    function MajDonneesSalariePop(Population : string; daterech : Tdatetime) : TStringList;
  private
    TobParametres, TobSalarie : TOB;
    ListSalForInClause : String;
    ListeParametresCharges : TStringList;
    ListeTypes : TStringList; //PT2
    Function GetCodePop(Population : String) : String;
    Procedure GetSalaries(TobListSalarie: TOB); overload;
    Procedure GetSalaries(StrListSalarie: TStringList); overload;
    {$IFDEF IDEMPOP}
    Procedure GetParametres(CodeParametre : String; Module : String = ''; TypeParametre : String = 'POP');
    procedure GetAssocSalPop;
    procedure GetAssocPopParam(daterech : Tdatetime);
    {$ENDIF}
    //Fonction OBSOLETE
    //Ancienne fonction de validation des modification
    //C'est maintenant le parametrage lui même qui est validé
//    Function ValideModifSalarie(TobSalarieAndUpdates: Tob) : Boolean;
  end;

Implementation
uses
  ENTPAIE,
  P5Def, StrUtils,
  {$IFDEF EAGLCLIENT}
    MaineAgl,
  {$ELSE}
    {$IFNDEF EAGLSERVER}
      FE_Main,                                   
    {$ENDIF}
  {$ENDIF}
   hmsgbox, hdtlinks, paramsoc, ComObj{$IFDEF PRESENCE}, PGPresence{$ENDIF};

//On stock la liste des type de population
//type TTableauTypePop = array[0..4] of String; //PT2
//const TableauTypePop : TTableauTypePop = ('PAI', 'IDR', 'HAB', 'FOR', 'PRE');
//On stock les associations
type TTableauAssociation = array[0..48, 0..2] of String;
                             //Module, Nom du champ type, Nom du champ
const TableauAssociation: TTableauAssociation =  (
                             ( 'PAI', 'PSA_TYPACPSOC'      , 'PSA_RIBACPSOC'    ),
                             ( 'PAI', 'PSA_TYPPAIACOMPT'   , 'PSA_PAIACOMPTE'   ),
                             ( 'PAI', 'PSA_TYPACTIVITE'    , 'PSA_ACTIVITE'     ),
                             ( 'PAI', 'PSA_TYPCONVENTION'  , 'PSA_CONVENTION'   ),
                             ( ''   , 'PSA_TYPDADSFRAC'    , 'PSA_DADSFRACTION' ),
                             ( 'PAI', 'PSA_TYPEDITBULCP'   , 'PSA_EDITBULCP'    ),
                             ( ''   , 'PSA_TYPEDITORG'     , 'PSA_EDITORG'      ),
                             ( 'PAI', 'PSA_TYPFRAISSOC'    , 'PSA_RIBFRAISSOC'  ),
                             ( ''   , 'PSA_TYPJOURHEURE'   , 'PSA_JOURHEURE'    ),
                             ( 'PAI', 'PSA_TYPNBACQUISCP'  , 'PSA_NBACQUISCP'   ),
                             ( 'PAI', 'PSA_CPACQUISMOIS'   , 'PSA_NBREACQUISCP' ),
                             ( 'PAI', 'PSA_CPACQUISSUPP'   , 'PSA_NBRECPSUPP'   ),
                             ( 'PAI', 'PSA_CPACQUISANC'    , 'PSA_BASANCCP'     ),
                             ( 'PAI', 'PSA_CPACQUISANC'    , 'PSA_VALANCCP'     ),
                             ( 'PAI', 'PSA_DATANC'         , 'PSA_DATEACQCPANC' ),
                             ( 'PAI', 'PSA_DATANC'         , 'PSA_TYPDATANC'    ),
                             ( 'PAI', 'PSA_TYPPAIEVALOMS'  , 'PSA_PAIEVALOMS'   ),
                             ( 'PAI', 'PSA_CPTYPEMETHOD'   , 'PSA_VALORINDEMCP' ),
                             ( 'PAI', 'PSA_CPTYPEVALO'     , 'PSA_MVALOMS'      ),
                             ( 'PAI', 'PSA_CPTYPEVALO'     , 'PSA_VALODXMN'     ),
                             ( 'PAI', 'PSA_TYPPAIFRAIS'    , 'PSA_PAIFRAIS'     ),
                             ( 'PAI', 'PSA_TYPPERIODEBUL'  , 'PSA_PERIODBUL'    ),
                             ( 'PAI', 'PSA_TYPPROFIL'      , 'PSA_PROFIL'       ),
                             ( 'PAI', 'PSA_TYPPROFILAFP'   , 'PSA_PROFILAFP'    ),
                             ( 'PAI', 'PSA_TYPPROFILANC'   , 'PSA_PROFILANCIEN' ),
                             ( 'PAI', 'PSA_TYPPROFILAPP'   , 'PSA_PROFILAPP'    ),
                             ( 'PAI', 'PSA_TYPPROFILCGE'   , 'PSA_PROFILCGE'    ),
                             ( ''   , 'PSA_TYPPROFILFNAL'  , 'PSA_PROFILFNAL'   ),
                             ( 'PAI', 'PSA_TYPPROFILMUT'   , 'PSA_PROFILMUT'    ),
                             ( 'PAI', 'PSA_TYPPROFILPRE'   , 'PSA_PROFILPRE'    ),
                             ( 'PAI', 'PSA_TYPPROFILRBS'   , 'PSA_PROFILRBS'    ),
                             ( 'PAI', 'PSA_TYPPROFILREM'   , 'PSA_PROFILREM'    ),
                             ( 'PAI', 'PSA_TYPPROFILRET'   , 'PSA_PROFILRET'    ),
                             ( ''   , 'PSA_TYPPROFILTPS'   , 'PSA_PROFILTPS'    ),
                             ( 'PAI', 'PSA_TYPPROFILTRANS' , 'PSA_PROFILTRANS'  ),
                             ( 'PAI', 'PSA_TYPPROFILTSS'   , 'PSA_PROFILTSS'    ),
                             ( ''   , 'PSA_TYPPRUDH'       , 'PSA_PRUDHCOLL'    ),
                             ( ''   , 'PSA_TYPPRUDH'       , 'PSA_PRUDHSECT'    ),
                             ( ''   , 'PSA_TYPPRUDH'       , 'PSA_PRUDHVOTE'    ),
                             ( 'PAI', 'PSA_TYPREDREPAS'    , 'PSA_REDREPAS'     ),
                             ( 'PAI', 'PSA_TYPREDRTT1'     , 'PSA_REDRTT1'      ),
                             ( 'PAI', 'PSA_TYPREDRTT2'     , 'PSA_REDRTT2'      ),
                             ( 'PAI', 'PSA_TYPREGLT'       , 'PSA_PGMODEREGLE'  ),
                             ( 'PAI', 'PSA_TYPDATPAIEMENT' , 'PSA_JOURPAIEMENT' ),
                             ( 'PAI', 'PSA_TYPDATPAIEMENT' , 'PSA_MOISPAIEMENT' ),
                             ( 'PAI', 'PSA_TYPVIRSOC'      , 'PSA_RIBVIRSOC'    ),
                             ( 'PAI', 'PSA_CPTYPERELIQ'    , 'PSA_RELIQUAT'     ),
                             ( ''   , 'PSA_TYPPERIODCT'    , 'PSA_RELIQUAT'     ),
                             ( ''   , 'PSA_STANDCALEND'    , '')
                           );

type TTableauPredefini = array[0..2] of String;
const TableauPredefini : TTableauPredefini = ('CEG', 'STD', 'DOS');

function FindTabletteName(IDENT : integer) : String;
var
  Q : TQuery;
  stQ : String;
begin
  try
    stQ := 'SELECT PAI_LIENASSOC FROM PAIEPARIM WHERE ##PAI_PREDEFINI## AND PAI_IDENT = '+intToStr(IDENT);
    Q := OpenSQL(stQ, True, 1);
    Result := Q.FindField('PAI_LIENASSOC').AsString;
    //La tablette PGLIBCOEFFICIENT indiquée dans PAIEPARIM n'est pas correcte,
    // on utilise PGCOEFFICIENT à la place
    if Result = 'PGLIBCOEFFICIENT' then Result := 'PGCOEFFICIENT';
    Ferme(Q);
  except
    result := '';
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 13/06/2007
Modifié le ... :   /  /    
Description .. : Renvoie le prédefini pris en compte par le paramétrage des 
Suite ........ : populations pour un type de populations donné
Mots clefs ... : 
*****************************************************************}
Function GetPredefiniPopulation(TypePop : String) : String;
var
  sql1 : String;
  TobPop : Tob;
begin
  //Optimisation de la recherche du prédéfini -> 1 seul requete
  sql1 := 'select PPC_POPULATION, PPO_CODEPOP, PPO_PREDEFINI,PPO_NODOSSIER from ORDREPOPULATION, CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
  ' and PPO_TYPEPOP LIKE "%'+ typepop +'%"';
  TobPop := Tob.Create('recherche du prédéfini',nil,-1);
  TobPop.LoadDetailFromSQL(sql1);
  result := 'DOS';
  if not Assigned(TobPop.FindFirst(['PPO_PREDEFINI','PPO_NODOSSIER'],['DOS',PgrendNodossier()],False)) then
  begin
    result := 'STD';
    if not Assigned(TobPop.FindFirst(['PPO_PREDEFINI'],['STD'],False)) then
    begin
      result := 'CEG';
      if not Assigned(TobPop.FindFirst(['PPO_PREDEFINI'],['CEG'],False)) then
      begin
        result := '';
      end;
    end;
  end;
  FreeAndNil(TobPop);
{  // recherche des codes pop prédéfini DOSSIER
  result := 'DOS';
  sql1 := 'select PPC_POPULATION, PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
  ' and PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+PgrendNodossier()+'" and PPO_TYPEPOP LIKE "%'+ typepop +'%"';
  if not existesql(sql1) then
  begin
    // recherche des codes pop prédéfini STANDARD
    result := 'STD';
    sql1 := 'select PPC_POPULATION, PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
    ' and PPO_PREDEFINI= "STD" and PPO_TYPEPOP LIKE "%'+ typepop +'%" ';
    if not existesql(sql1) then
    begin
      result := 'CEG';
      sql1 := 'select PPC_POPULATION, PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
      ' and PPO_PREDEFINI= "CEG" and PPO_TYPEPOP LIKE "%'+ typepop +'%" ';
      if not existesql(sql1) then
      begin
        result := '';
      end;
    end;
  end; }
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 13/06/2007
Modifié le ... :   /  /    
Description .. : Renvoie la liste des champs utilisés dans le paramétrage 
Suite ........ : des populations
Mots clefs ... : 
*****************************************************************}
Function GetFieldsListFromPopulParam(TypePop : String) : TStringList;
var
  sql1, predefini : String;
  Q1: TQuery;
  CriteresList : String;
  indexident : integer;
begin
  result := TStringList.Create;
  predefini := GetPredefiniPopulation(TypePop);
  // Deb PT6
  // Si IDR non sérialisé, faire la requête pour les types de pop <> de IDR
  if Typepop = '' then
  begin
    if  (VH_Paie.PgSeriaIDR = false) then
      sql1 := 'SELECT ppo_nbident, ppo_ident1, ppo_ident2, ppo_ident3, ppo_ident4 FROM CODEPOPULATION '
            + ' where ##ppo_predefini## and ppo_predefini = "'+predefini+'" AND PPO_TYPEPOP not like "%IDR%"'
    else
      sql1 := 'SELECT ppo_nbident, ppo_ident1, ppo_ident2, ppo_ident3, ppo_ident4 FROM CODEPOPULATION '
        + ' where ##ppo_predefini## and ppo_predefini = "'+predefini+'" ';
  end
  else  // Typepop <> ''
      sql1 := 'SELECT ppo_nbident, ppo_ident1, ppo_ident2, ppo_ident3, ppo_ident4 FROM CODEPOPULATION '
        + ' where ##ppo_predefini## and ppo_predefini = "'+predefini+'" and ppo_typepop like "%'+TypePop+'%" ';

  // fin pt6
  //On récupère la liste des champs utilisés comme critères
  // pt6 sql1 := 'SELECT ppo_nbident, ppo_ident1, ppo_ident2, ppo_ident3, ppo_ident4 FROM CODEPOPULATION '
  //pt6       + ' where ##ppo_predefini## and ppo_predefini = "'+predefini+'" ';
  //pt6  if TypePop <> '' then sql1 := sql1 + ' and ppo_typepop like "%'+TypePop+'%"';

  Q1 := opensql(sql1, true);
  CriteresList := '';
  while not Q1.EOF do
  begin
    for indexident := 1 to Q1.Fields[0].AsInteger do
    begin
      CriteresList := CriteresList + Q1.Fields[indexident].AsString +', ';
    end;
    Q1.Next;
  end;
  ferme(Q1);
  if CriteresList = '' then exit;
  CriteresList := leftstr(CriteresList, length(CriteresList)-2);
  //On récupère les champs correspondants
  sql1 := 'SELECT pai_prefix, pai_suffix FROM PAIEPARIM where pai_ident in ('+CriteresList+')';
  Q1 := opensql(sql1, true);
  result.Clear;
  while not Q1.EOF do
  begin
    result.Add(Q1.Fields[0].AsString+'_'+Q1.Fields[1].AsString);
    Q1.Next;
  end;
  ferme(Q1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 13/06/2007
Modifié le ... :   /  /    
Description .. :  Vrai si l'un des champs passé en tableau fait parti des 
Suite ........ : champs utilisés dans le paramétrage des populations
Mots clefs ... : 
*****************************************************************}
Function isIncludeInPopulParam(ListChamps : Array of String; TypePop : String = '') : Boolean;
var
  indexChamp : Integer;
  ListChampsFromParam : TStringList;
begin
  result := False;
  //On récupère les champs concernés par le paramétrage de ce type de population
  ListChampsFromParam := GetFieldsListFromPopulParam(typepop);
  //On le compare à la liste de champs modifiés
  for indexChamp := 0 to length(ListChamps) do
  begin
    //si au moins un champ est identique, on lance la mise à jour
    if ListChampsFromParam.IndexOf(ListChamps[indexChamp]) > -1 then
    begin
      result := True;
      exit;
    end;
  end;
  FreeAndNil(ListChampsFromParam);
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 20/06/2006
Modifié le ... : 20/06/2006
Description .. : Fonction : Constitution de la TOB salarié pour une
Suite ........ : population
Mots clefs ... :
*****************************************************************}
function salariepop(Codepop, population : string; datesimul : Tdatetime) :  TOB;
var
  st1,st2, sql1 ,cassql, finper: string;
  TOB_SAL, T1 : TOB;
  histo, paieencours : boolean;
  Q , Q1: TQuery;
  finperdate : TDatetime;
begin
  paieencours := false;
 // Q1 := opensql('SELECT PEX_FINPERIODE FROM EXERSOCIAL WHERE PEX_ACTIF = "X" ORDER BY PEX_ANNEEREFER DESC', true);   // pt3
  Q1 := opensql('SELECT PEX_DEBUTPERIODE FROM EXERSOCIAL WHERE PEX_ACTIF = "X" ORDER BY PEX_ANNEEREFER DESC', true);   // pt3
  if not Q1.EOF then
  begin
   // finper := Q1.findfield('PEX_FINPERIODE').asstring;   // pt3
    finper := Q1.findfield('PEX_DEBUTPERIODE').asstring;     // pt3
    finperdate := StrtoDate(finper);
    if datesimul >= finperdate then paieencours := false else paieencours := true;
  end;
  ferme(Q1);

  histo := false;
  if (paieencours and vh_paie.pghistorisation) then histo := true;

  TOB_SAL := Tob.Create('Les salariés',Nil,-1);

  // recherche de la clause Where pour cette population
  PGrechwhere(Codepop, population, histo, st1,st2);

  // recherche des donnes SALARIES avec la clause where SALARIES
  if (histo = false) or (st2 = '') then
  begin
  sql1 := 'select PSA_SALARIE, PSA_LIBELLE,PSA_PRENOM,PSA_NUMEROSS,PSA_SEXE,PSA_DATENAISSANCE,PSA_DATEENTREE,'+
  ' PSA_DATEANCIENNETE,PSA_CONVENTION,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_LIBREPCMB1,'+
  ' PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_CODESTAT,PSA_DADSPROF,PSA_LIBELLEEMPLOI,pSA_CONDEMPLOI,'+
  ' PSA_DADSCAT,PSA_ETABLISSEMENT from SALARIES where '+st1+' and (PSA_DATESORTIE >="'+usdatetime(datesimul)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+
  usdatetime(datesimul)+'")';
  cassql := 'S';
  end
  else
   begin
   cassql := 'H';
   if st1 <> '' then
  sql1 := 'Select PSA_SALARIE, PSA_LIBELLE,PSA_PRENOM,PSA_NUMEROSS,PSA_SEXE,PSA_DATENAISSANCE,PSA_DATEENTREE,'+
  ' PSA_DATEANCIENNETE,PSA_CONVENTION,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_LIBREPCMB1,'+
  ' PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_CODESTAT,PSA_DADSPROF,PSA_LIBELLEEMPLOI,pSA_CONDEMPLOI,'+
  ' PSA_DADSCAT,PSA_ETABLISSEMENT,PHS_TRAVAILN1,PHS_TRAVAILN2,PHS_TRAVAILN3,PHS_TRAVAILN4,PHS_CBLIBRE1,PHS_CBLIBRE2,'+
  ' PHS_CBLIBRE3,PHS_CBLIBRE4,PHS_CODESTAT,PHS_DADSPROF,PHS_LIBELLEEMPLOI,PHS_CONDEMPLOI,PHS_DADSCAT,PHS_ETABLISSEMENT '+
  ' from SALARIES,  HISTOSALARIE a where '+st1+' and (PSA_DATESORTIE >="'+usdatetime(datesimul)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <='+
  ' "'+usdatetime(datesimul)+'") and psa_salarie = phs_salarie and a.Phs_dateapplic ='+
  '(select max(phs_dateapplic) from HISTOSALARIE b Where b.phs_dateapplic<= "'+UsdateTime(datesimul)+'"'+
  ' and a.phs_salarie= b.phs_salarie) and '+st2+''
  else
  sql1 := 'Select PSA_SALARIE, PSA_LIBELLE,PSA_PRENOM,PSA_NUMEROSS,PSA_SEXE,PSA_DATENAISSANCE,PSA_DATEENTREE,'+
  ' PSA_DATEANCIENNETE,PSA_CONVENTION,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_LIBREPCMB1,'+
  ' PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_CODESTAT,PSA_DADSPROF,PSA_LIBELLEEMPLOI,pSA_CONDEMPLOI,'+
  ' PSA_DADSCAT,PSA_ETABLISSEMENT,PHS_TRAVAILN1,PHS_TRAVAILN2,PHS_TRAVAILN3,PHS_TRAVAILN4,PHS_CBLIBRE1,PHS_CBLIBRE2,'+
  ' PHS_CBLIBRE3,PHS_CBLIBRE4,PHS_CODESTAT,PHS_DADSPROF,PHS_LIBELLEEMPLOI,PHS_CONDEMPLOI,PHS_DADSCAT,PHS_ETABLISSEMENT '+
  ' from SALARIES,  HISTOSALARIE a where (PSA_DATESORTIE >="'+usdatetime(datesimul)+'" or'+
  ' PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <='+
  ' "'+usdatetime(datesimul)+'") and psa_salarie = phs_salarie and a.Phs_dateapplic ='+
  '(select max(phs_dateapplic) from HISTOSALARIE b Where b.phs_dateapplic<= "'+UsdateTime(datesimul)+'"'+
  ' and a.phs_salarie= b.phs_salarie) and '+st2+'';
  end;

  Q := Opensql(sql1, true);

  // Constitution de la Tob salarié avec les données venant de SALARIES ou HISTOSALARIE
  while not Q.EOF do
  begin
    T1 := Tob.Create('Salarié',TOB_SAL,-1);
    T1.AddChampSupValeur('SALARIE',Q.findfield('PSA_SALARIE').asstring);
    T1.AddChampSupValeur('NOM',Q.findfield('PSA_LIBELLE').asstring);
    T1.AddChampSupValeur('PRENOM',Q.findfield('PSA_PRENOM').asstring);
    T1.AddChampSupValeur('NUMEROSS',Q.findfield('PSA_NUMEROSS').asstring);
    T1.AddChampSupValeur('SEXE',Q.findfield('PSA_SEXE').asstring);
    T1.AddChampSupValeur('DATENAISS',Q.findfield('PSA_DATENAISSANCE').asstring);
    T1.AddChampSupValeur('DATEENTREE',Q.findfield('PSA_DATEENTREE').asstring);
    T1.AddChampSupValeur('DATEANCIENNETE',Q.findfield('PSA_DATEANCIENNETE').asstring);
    T1.AddChampSupValeur('CONVENTION',Q.findfield('PSA_CONVENTION').asstring);
    T1.AddChampSup('TRAVAILN1',False);
    T1.AddChampSup('TRAVAILN2',False);
    T1.AddChampSup('TRAVAILN3',False);
    T1.AddChampSup('TRAVAILN4',False);
    T1.AddChampSup('LIBREPCMB1',False);
    T1.AddChampSup('LIBREPCMB2',False);
    T1.AddChampSup('LIBREPCMB3',False);
    T1.AddChampSup('LIBREPCMB4',False);
    T1.AddChampSup('CODESTAT',False);
    T1.AddChampSup('DADSPROF',False);
    T1.AddChampSup('LIBELLEEMPLOI',False);
    T1.AddChampSup('CONDEMPLOI',False);
    T1.AddChampSup('DADSCAT',False);
    T1.AddChampSup('ETABL',False);
    if cassql = 'S' then
    begin
    // si les données proviennent de SALARIES
    T1.putValue('TRAVAILN1',Q.findfield('PSA_TRAVAILN1').asstring);
    T1.putValue('TRAVAILN2',Q.findfield('PSA_TRAVAILN2').asstring);
    T1.putValue('TRAVAILN3',Q.findfield('PSA_TRAVAILN3').asstring);
    T1.putValue('TRAVAILN4',Q.findfield('PSA_TRAVAILN4').asstring);
    T1.putValue('LIBREPCMB1',Q.findfield('PSA_LIBREPCMB1').asstring);
    T1.putValue('LIBREPCMB2',Q.findfield('PSA_LIBREPCMB2').asstring);
    T1.putValue('LIBREPCMB3',Q.findfield('PSA_LIBREPCMB3').asstring);
    T1.putValue('LIBREPCMB4',Q.findfield('PSA_LIBREPCMB4').asstring);
    T1.putValue('CODESTAT',Q.findfield('PSA_CODESTAT').asstring);
    T1.putValue('DADSPROF',Q.findfield('PSA_DADSPROF').asstring);
    T1.putValue('LIBELLEEMPLOI',Q.findfield('PSA_LIBELLEEMPLOI').asstring);
    T1.putValue('CONDEMPLOI',Q.findfield('PSA_CONDEMPLOI').asstring);
    T1.putValue('DADSCAT',Q.findfield('PSA_DADSCAT').asstring);
    T1.PutValue('ETABL',Q.findfield('PSA_ETABLISSEMENT').asstring);
    end
    else
    begin
    // si les données proviennent de HISTOSALARIE
    T1.putValue('TRAVAILN1',Q.findfield('PHS_TRAVAILN1').asstring);
    T1.putValue('TRAVAILN2',Q.findfield('PHS_TRAVAILN2').asstring);
    T1.putValue('TRAVAILN3',Q.findfield('PHS_TRAVAILN3').asstring);
    T1.putValue('TRAVAILN4',Q.findfield('PHS_TRAVAILN4').asstring);
    T1.putValue('LIBREPCMB1',Q.findfield('PHS_CBLIBRE1').asstring);
    T1.putValue('LIBREPCMB2',Q.findfield('PHS_CBLIBRE2').asstring);
    T1.putValue('LIBREPCMB3',Q.findfield('PHS_CBLIBRE3').asstring);
    T1.putValue('LIBREPCMB4',Q.findfield('PHS_CBLIBRE4').asstring);
    T1.putValue('CODESTAT',Q.findfield('PHS_CODESTAT').asstring);
    T1.putValue('DADSPROF',Q.findfield('PHS_DADSPROF').asstring);
    T1.putValue('LIBELLEEMPLOI',Q.findfield('PHS_LIBELLEEMPLOI').asstring);
    T1.putValue('CONDEMPLOI',Q.findfield('PHS_CONDEMPLOI').asstring);
    T1.putValue('DADSCAT',Q.findfield('PHS_DADSCAT').asstring);
    T1.PutValue('ETABL',Q.findfield('PHS_ETABLISSEMENT').asstring);
    end;
    Q.Next;
  end;
  ferme(Q);
  result := TOB_SAL;
//  TOB.Create('Liste des salariés',nil,-1);
//  result.Dupliquer(TOB_SAL,True,True);
//  FreeAndNil(TOB_SAL)
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 13/06/2007
Modifié le ... :   /  /    
Description .. : Fonction a utiliser pour valider ou dévalider un type de 
Suite ........ : population (notemment en cas de changements)
Mots clefs ... : 
*****************************************************************}
Procedure SetValidePopulParamSoc(TypePop : String; Valide : Boolean);
var
  ListeTypePop, NewListeTypePop, TempTypePop : String;
  isAlreadyValid : Boolean;
begin
  ListeTypePop := VH_Paie.PGPopulValides;
  isAlreadyValid := False;
  TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
  While (ListeTypePop <> '') or (TempTypePop <> '') do
  begin
    if CompareText(TempTypePop,TypePop) = 0 then
      isAlreadyValid := True;
    TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
  end;
  if Valide and isAlreadyValid then exit;
  if (not Valide) and (not isAlreadyValid) then exit;
  if (Valide) and (not isAlreadyValid) then
  begin
    NewListeTypePop := VH_Paie.PGPopulValides+TypePop+';';
  end;
  if (not Valide) and (isAlreadyValid) then
  begin
    ExecuteSQL('Delete from SALARIEPOPUL where PNA_TYPEPOP ="'+typepop+'"');
    NewListeTypePop := '';
    ListeTypePop := VH_Paie.PGPopulValides;
    TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
    While (ListeTypePop <> '') or (TempTypePop <> '') do
    begin
      if CompareText(TempTypePop,TypePop) <> 0 then
        NewListeTypePop := NewListeTypePop+TempTypePop+';';
      TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
    end;
  end;
  SetParamSoc('SO_PGPOPULVALIDES',NewListeTypePop);
  VH_Paie.PGPopulValides := NewListeTypePop;
{$IFDEF EAGLCLIENT}
  AvertirCacheServer('PARAMSOC');
{$ENDIF}
end;


Procedure GetTablettePrefixeRetour(TabletteName : String; var PrefixeTablette, RetourTablette : String);
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT DO_CODE, DO_PREFIXE FROM DECOMBOS WHERE DO_COMBO="'+TabletteName+'"', True, 1);
  PrefixeTablette := Q.FindField('DO_PREFIXE').AsString;
  RetourTablette := Q.FindField('DO_CODE').AsString;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Fonction qui va vérifier dans les paramètres société que la
Suite ........ : population a été validée et peut être utilisée.
Mots clefs ... :
*****************************************************************}
Function CheckValidPopulParametres(TypePop : String = '') : Boolean;
var
  ListPopValides, tempPopValide : String;
begin
  result := False;
  ListPopValides := VH_Paie.PGPopulValides;
  tempPopValide := Trim(ReadTokenPipe(ListPopValides, ';'));
  While (ListPopValides <> '') or (tempPopValide <> '') do
  begin
    if CompareText(tempPopValide,TypePop) = 0 then
      result := True;
    tempPopValide := Trim(ReadTokenPipe(ListPopValides, ';'));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 26/04/2007
Modifié le ... :   /  /    
Description .. : Fonction a appeler avant d'utiliser les populations.
Suite ........ : L'utilisation des populations n'est permise que si la fonction 
Suite ........ : renvoi True
Mots clefs ... : 
*****************************************************************}
Function CanUsePopulation(TypePop : String = '') : Boolean;
var
{$IFNDEF EAGLSERVER}
  stModalResult : String;
{$ENDIF}
  Predef : String;
  stTempTypePop, stTypePop : String;
//  Q : TQuery;
{$IFNDEF EAGLSERVER}
  stType : String;
{$ELSE}
  UpdateIdemPop : TUpdateIdemPop;
{$ENDIF}
begin
  result := CheckValidPopulParametres(TypePop);
  if result = False then
  begin
    stTypePop := TypePop;
    stTempTypePop := ReadTokenSt(stTypePop);
    While (stTempTypePop <> '') do
    Begin
      Predef := Predef + GetPredefiniPopulation(stTempTypePop);
      stTempTypePop := ReadTokenSt(stTypePop);
    end;
    if Predef = '' then exit;
    // Les populations ne sont pas actives pour ce type

    if ((typepop = 'IDR') and (VH_Paie.PgSeriaIDR = false)) then exit;   // pt6

    // On lance la fenêtre de validation
{$IFNDEF EAGLSERVER}
    if TypePop = '' then
      stType := ''
    else
      stType := TraduireMemoire('de type ')+TypePop;
    if PGIAsk(TraduireMemoire('Le paramétrage des populations ')+stType+TraduireMemoire(' n''est pas valide.#10#13Voulez-vous le valider ?'),TraduireMemoire('Gestion des populations')) = mrYes then
    begin
      stModalResult := '';
      stModalResult := AGLLanceFiche('PAY', 'AFFECTSALPOPUL', '', '',TypePop);
      if stModalResult = 'True' then
      begin
        result := True;
      end;
      if not result then PGIError(TraduireMemoire('La validation du paramétrage des populations de type ')+TypePop+TraduireMemoire(' a échoué.'),TraduireMemoire('Gestion des populations'));
    end;
{$ELSE}
    if ValidPopulParametres(TypePop) then
    begin
      // Appel de la procédure d'affectation
      stTypePop := TypePop;
      stTempTypePop := ReadTokenSt(stTypePop);
      While (stTempTypePop <> '') do
      Begin
        UpdateIdemPop := TUpdateIdemPop.Create;
        UpdateIdemPop.MajSALARIEPOPUL(stTempTypePop,date());
        UpdateIdemPop.Free;
        stTempTypePop := ReadTokenSt(stTypePop);
      end;
      result := True;
    end;
{$ENDIF}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Contrôle de l'unicité et de l'exhaustivité du paramétrage des
Suite ........ : populations.
Mots clefs ... :
*****************************************************************}
Function ValidPopulParametres(TypePop : String = '') : Boolean;
var
  indexIdent, indexTypePop : Integer;
  QPopulations, QCriteres : TQuery;
  TobModules, TobModule, TobPopulation : TOB;
  TobModulesCritere, TobCritere, TempTobModulesCritereTYPEPOP : TOB;
  ParamPopulValidator : TParamPopulValidator;
  indexModulesCritere, indexModule : integer;
  TempTypePop : String;
  TempPredefini : String;
  TempMatrice : TMatriceCompteurReference;
  ListeTypePop : String;
  ListeTypes : TStringList;
begin
  ListeTypes := GetTypesPop; //PT2

  if (TypePop = '') or (TypePop = '<<Tous>>') then
  begin
    TypePop := '';
    for indexModule := 0 to ListeTypes.Count-1 do //PT2
    begin
      TypePop := TypePop + ListeTypes[indexModule] +';'; //PT2
    end;
  end;

//Création de la tob des différents types de population
  TobModules := Tob.create('type de populations',Nil,-1);
  for indexTypePop := 0 to ListeTypes.Count-1 do //PT2
  begin
    if pos(ListeTypes[indexTypePop],TypePop) > 0 then //PT2
    begin
      TobModule := Tob.create('type de populations',TobModules,-1);
      TobModule.AddChampSupValeur('TYPEPOP',ListeTypes[indexTypePop]); //PT2
    end;
  end;
  TobModulesCritere := Tob.create('type de populations',Nil,-1);
  TobModulesCritere.Dupliquer(TobModules,True,True);
  QPopulations := opensql( ' SELECT PPC_PREDEFINI, PPC_POPULATION, PPC_CODEPOP as PGO_CODEASSOCIE, PPO_TYPEPOP FROM ORDREPOPULATION '
                          +' left outer join codepopulation on PPC_CODEPOP = PPO_CODEPOP '
                          +' WHERE ##PPC_PREDEFINI## '// AND PPO_TYPEPOP like "%IDR%" '
                          ,True);
  //Parcours de chaque type de population
  for indexTypePop := 0 to ListeTypes.Count-1 do //PT2
  begin
    if pos(ListeTypes[indexTypePop],TypePop) > 0 then //PT2
    begin
      //Recherche de la tob correspondant au type
      TobModule := TobModules.FindFirst(['TYPEPOP'],[ListeTypes[indexTypePop]],False); //PT2
      if TobModule <> nil then
      begin
        while not QPopulations.Eof do
        begin
          //Ajout de la population
          if pos(ListeTypes[indexTypePop], QPopulations.Fields[3].AsString) > 0 then //PT2
          begin
            TobPopulation := Tob.create('Population',TobModule,-1);
            TobPopulation.AddChampSupValeur('PGO_CODEASSOCIE',QPopulations.Fields[2].AsString);
            TobPopulation.AddChampSupValeur('PGO_PREDEFINI',QPopulations.Fields[0].AsString);
          end;
          QPopulations.Next;
        end;
        QPopulations.First;
      end;
    end;
  end;
  ferme(QPopulations);
//Chargement des critères de chaque population
  QCriteres := opensql('SELECT PPO_CODEPOP, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4, PPO_PREDEFINI FROM CODEPOPULATION WHERE ##PPO_PREDEFINI##',True);
  while not QCriteres.Eof do
  begin
    //Recherche des tobs correspondant à la population
    TobPopulation := TobModules.FindFirst(['PGO_CODEASSOCIE'],[QCriteres.Fields[0].AsString],True);
    While Assigned(TobPopulation) do
    begin
      TempTypepop := TobPopulation.Parent.GetString('TYPEPOP');
      TempTobModulesCritereTYPEPOP := TobModulesCritere.FindFirst(['TYPEPOP'],[TempTypepop],False);
      //Ajout du critère
      for indexIdent := 1 to 4 do
      begin
        if QCriteres.Fields[indexIdent].AsInteger <> 0 then
        begin
          if not Assigned(TempTobModulesCritereTYPEPOP.FindFirst(['PPO_IDENT'],[QCriteres.Fields[indexIdent].AsString],True)) then
          begin
            TobCritere := Tob.create('Critère' ,TempTobModulesCritereTYPEPOP ,-1);
            TobCritere.AddChampSupValeur('PPO_IDENT',QCriteres.Fields[indexIdent].AsString);
            TobCritere.AddChampSupValeur('PPO_PREDEFINI',QCriteres.Fields[5].AsString);
          end;
        end;
      end;
      TobPopulation := TobModules.FindNext(['PGO_CODEASSOCIE'],[QCriteres.Fields[0].AsString],True);
    end;
    QCriteres.Next;
  end;
  ferme(QCriteres);
//On récupère les différentes valeurs possibles pour chacun de ces critères
//On construit une matrice de compteur de références pour les critères
  ParamPopulValidator := TParamPopulValidator.Create;
  for indexModulesCritere := 0 to TobModulesCritere.FillesCount(0) -1 do
  begin
    TempTypePop := TobModulesCritere.Detail[indexModulesCritere].GetString('TYPEPOP');
    for indexModule := 0 to TobModulesCritere.Detail[indexModulesCritere].FillesCount(1) -1 do
    begin
      TempPredefini := TobModulesCritere.Detail[indexModulesCritere].Detail[indexModule].GetString('PPO_PREDEFINI');
      TempMatrice := ParamPopulValidator.GetMatrice(TempTypePop, TempPredefini);
      if TempMatrice <> nil then
        TempMatrice.AddCritere(TobModulesCritere.Detail[indexModulesCritere].Detail[indexModule].GetInteger('PPO_IDENT'));
    end;
  end;
  FreeAndNil(TobModules);
  FreeAndNil(TobModulesCritere);
//On rempli la matrice en comptant les références
  ParamPopulValidator.UpdateCompteurs;
//On vérifie que chaque compteur de référence est unitaire
  result := True;
  ListeTypePop := TypePop;
  TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
  While (ListeTypePop <> '') or (TempTypePop <> '') do
  begin
    if not ParamPopulValidator.IsValide(TempTypePop) then
    begin
      SetValidePopulParamSoc(TempTypePop,False);
      result := False;
    end else begin
      SetValidePopulParamSoc(TempTypePop,True);
    end;
    TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
  end;
  ParamPopulValidator.Free;
  ListeTypes.Free; //PT2
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Recherche de la population d'un salarié selon le type de pop
                 avec en paramètre :
                 Salarié : salarié à traiter
                 Type population : type de la population
                 Résultat : population      
Mots clefs ... :
*****************************************************************}
function pgpopsalarie(salarie, typepop : string): string;
var
  sql1,sql2,sql3,sql4,st1, nodossier, population, codepop, populsalarie, sti, numident,nomchamp, pred: String;
  Q1, QQ : TQuery; //, Q0, Q3
  tob_Identsal, Tids , tob_valident , tob_paieparim , TD: TOB;
  i,  nbident: integer;
  trouve : boolean;
  Tob_TempValIdent, Tob_TempParentValIdent, Tob_TempFilleValIdent : TOB;
begin
  // Recherche des champs de PAIEPARIM utilisable comme critères de population
//  Q0 := Opensql('select PAI_IDENT,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_UTILISABLEPOP <> ""', true);
  Tob_paieparim := Tob.Create('Critères population', nil, -1);
//  Tob_paieparim.LoadDetailDB('Critères population', '', '', Q0, False);
  Tob_paieparim.LoadDetailFromSQL('select PAI_IDENT,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_UTILISABLEPOP <> ""');
//  ferme(Q0);
  trouve := false;
  populsalarie := '';
  // lecture des codes critères concernés par le type de population
  nodossier := PgrendNodossier();
  pred := GetPredefiniPopulation(typepop);
  if pred = 'DOS' then
  sql1 :='select PPO_CODEPOP from CODEPOPULATION where PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and'+
  ' PPO_TYPEPOP LIKE "%'+ typepop +'%"'
  else
  sql1 := 'select PPO_CODEPOP from CODEPOPULATION where PPO_PREDEFINI= "'+pred+'" and PPO_TYPEPOP LIKE "%'+ typepop +'%"' ;
  QQ:= Opensql(sql1, true);
  While Not QQ.Eof Do
  begin   // d1
    codepop := QQ.findfield('PPO_CODEPOP').asstring;
    // recherche des identifiants
    sql2 := 'SELECT PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP FROM CODEPOPULATION' +
    ' Where PPO_CODEPOP = "'+codepop+'" ';
    Q1 := OPENSQL(sql2, True);
    if not Q1.Eof then
    begin
      nbident := Q1.findfield('PPO_NBIDENT').asinteger;
      if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
      Tob_Identsal := Tob.Create('Nom identifiants salarie', nil, -1);
      for i := 1 to nbident do
      begin  // d2
        Sti := IntToStr(i);
        numident := Q1.findfield('PPO_IDENT' + sti).asstring;
        TD := Tob_paieparim.findfirst(['PAI_IDENT'], [numident], true);
        nomchamp := TD.getvalue('PAI_PREFIX') + '_' + TD.getvalue('PAI_SUFFIX');
        Tids := Tob.Create('Identifiant', Tob_Identsal, -1);
        Tids.AddChampSup('NOMCHAMPS', False);
        Tids.PutValue('NOMCHAMPS', Nomchamp);
        Tids.AddChampSup('ORDREIDENT',False);
        Tids.PutValue('ORDREIDENT',sti);
      end; // e2
      // Lecture des populations pour ce code population
      sql3 := 'SELECT PPC_POPULATION, PPC_LIBELLE,PPC_VALIDENT1,PPC_VALIDENT2, PPC_VALIDENT3, PPC_VALIDENT4'+
      ' FROM ORDREPOPULATION WHERE ##PPC_PREDEFINI## PPC_CODEPOP = "' + codepop + '" ';
//      Q3 := OPENSQL(sql3, True);
      if assigned(Tob_valident) then FreeAndNil(Tob_valident);
      Tob_ValIdent := Tob.Create('Valeur identifiants', nil, -1);
//      Tob_ValIdent.LoadDetailDB('Valeur identifiants', '', '', Q3, False);
      Tob_ValIdent.LoadDetailFromSQL(sql3);
//      ferme(Q3);
      Tob_TempValIdent := Tob_ValIdent.FindFirst([''],[''],False);
      While Tob_TempValIdent <> nil do
      begin   // d3
        population := Tob_TempValIdent.GetString('PPC_POPULATION'); // Q3.findfield('PPC_POPULATION').asstring;
        // constitution de la clause WHERE
        st1 := '';
        //On crée une tob parente pour la tob temporaire, car c'est la structure
        // que la fonction PGConstitutionwhere attends
        Tob_TempParentValIdent := Tob.Create('Tob mère de Valeur identifiants', nil, -1);
        Tob_TempFilleValIdent :=  Tob.Create('Tob fille de Valeur identifiants', Tob_TempParentValIdent, -1);
        Tob_TempFilleValIdent.Dupliquer(Tob_TempValIdent, False, True);
        if Assigned(tob_identsal) then
          st1 := PGConstitutionwhere(Tob_identsal, Tob_TempParentValIdent);//Tob_Valident);
        // On libère la tob
        FreeAndNil(Tob_TempParentValIdent);
        if st1 <> '' then
          sql4 := 'select PSA_SALARIE from SALARIES where PSA_SALARIE = "'+salarie+'"  and '+st1+''
        else
          sql4 := 'select PSA_SALARIE from SALARIES where PSA_SALARIE = "'+salarie+'"';
        if  existesql(sql4) then
        begin
          populsalarie := population;
          trouve := true;
          break;
        end;
        Tob_TempValIdent := Tob_ValIdent.FindNext([''],[''],False);
      end; // e3
    end;
    ferme(Q1);
    if trouve then break;
    QQ.next;
  end; // e1
  ferme(QQ);
  if assigned(Tob_valident) then FreeAndNil(Tob_valident);
  if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
  if assigned(Tob_Paieparim) then FreeAndNil(Tob_Paieparim);
  result:= populsalarie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 23/02/2007
Modifié le ... :   /  /
Description .. : Liste des salariés et la population pour un type de
Suite ........ : population  
Mots clefs ... : 
*****************************************************************}
function pgsalarietypepop(typepop,predefini : string;  daterech:Tdatetime): TOB;
var
  sql1, sql3, sqlsal, codepop,nomchamp, population,libellepop, st1, sti, numident, nodossier: string;
  Q4, QQ: Tquery; //Q0, Q3,
  tob_Identsal, Tids , tob_valident, tob_saltypepop, T1 ,tob_paieparim , TD: TOB;
  i,  nbident: integer;
  Tob_TempValIdent, Tob_TempParentValIdent, Tob_TempFilleValIdent : TOB;
begin
  nodossier := pgrendNodossier();
  Tob_SALtypepop := Tob.Create('salariés/population',Nil,-1);
  // Recherche des champs de PAIEPARIM utilisable comme critères de population
//  Q0 := Opensql('select PAI_IDENT,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_UTILISABLEPOP <> ""', true);
  Tob_paieparim := Tob.Create('Critères population', nil, -1);
//  Tob_paieparim.LoadDetailDB('Critères population', '', '', Q0, False);
  Tob_paieparim.LoadDetailFromSQL('select PAI_IDENT,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_UTILISABLEPOP <> ""');
//  ferme(Q0);
  // lecture des codes critères concernés par le type de population
  if predefini = 'DOS' then
    sql1 := 'select PPO_CODEPOP, PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP from CODEPOPULATION where PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and'+
    ' PPO_TYPEPOP LIKE "%'+ typepop +'%" '
  else
    sql1 := 'select PPO_CODEPOP, PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP from CODEPOPULATION where PPO_PREDEFINI= "'+predefini+'" and PPO_TYPEPOP LIKE "%'+ typepop +'%"';
  QQ:= Opensql(sql1, true);
  While Not QQ.Eof Do
  begin   // d1
    codepop := QQ.findfield('PPO_CODEPOP').asstring;
    // recherche des identifiants dans PAIPARIM
    nbident := QQ.findfield('PPO_NBIDENT').asinteger;
    if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
    Tob_Identsal := Tob.Create('Nom identifiants salarie', nil, -1);
    for i := 1 to nbident do
    begin  // d2
      Sti := IntToStr(i);
      numident := QQ.findfield('PPO_IDENT' + sti).asstring;
      TD := Tob_paieparim.findfirst(['PAI_IDENT'], [numident], true);
      nomchamp := TD.getvalue('PAI_PREFIX') + '_' + TD.getvalue('PAI_SUFFIX');
      Tids := Tob.Create('Identifiant', Tob_Identsal, -1);
      Tids.AddChampSup('NOMCHAMPS', False);
      Tids.PutValue('NOMCHAMPS', Nomchamp);
      Tids.AddChampSup('ORDREIDENT',False);
      Tids.PutValue('ORDREIDENT',sti);
    end; // e2
   // Lecture des populations pour ce code population
    sql3 := 'SELECT PPC_POPULATION, PPC_LIBELLE,PPC_VALIDENT1,PPC_VALIDENT2, PPC_VALIDENT3, PPC_VALIDENT4'+
    ' FROM ORDREPOPULATION WHERE ##PPC_PREDEFINI## PPC_CODEPOP = "' + codepop + '" ';
//    Q3 := OPENSQL(sql3, True);
    if assigned(Tob_valident) then FreeAndNil(Tob_valident);
    Tob_ValIdent := Tob.Create('Valeur identifiants', nil, -1);
//    Tob_ValIdent.LoadDetailDB('Valeur identifiants', '', '', Q3, False);
    Tob_ValIdent.LoadDetailFromSQL(sql3);
//    ferme(Q3);
    Tob_TempValIdent := Tob_ValIdent.FindFirst([''],[''],False);
    While Tob_TempValIdent <> nil do //Not Q3.Eof Do
    begin   // d3
      population := Tob_TempValIdent.GetString('PPC_POPULATION'); // Q3.findfield('PPC_POPULATION').asstring;
      libellepop := Tob_TempValIdent.GetString('PPC_LIBELLE'); //Q3.findfield('PPC_LIBELLE').asstring;
      // constitution de la clause WHERE
      st1 := '';
      //On crée une tob parente pour la tob temporaire, car c'est la structure
      // que la fonction PGConstitutionwhere attends
      Tob_TempParentValIdent := Tob.Create('Tob mère de Valeur identifiants', nil, -1);
      Tob_TempFilleValIdent :=  Tob.Create('Tob fille de Valeur identifiants', Tob_TempParentValIdent, -1);
      Tob_TempFilleValIdent.Dupliquer(Tob_TempValIdent, False, True);
      if Assigned(tob_identsal) then
        st1 := PGConstitutionwhere(Tob_identsal, Tob_TempParentValIdent);//Tob_Valident);
      // On libère la tob
      FreeAndNil(Tob_TempParentValIdent);
      // recherche des salariés avec cette clause where
      if st1 <> '' then
        sqlsal := 'select PSA_SALARIE, PSA_PRENOM, PSA_LIBELLE from SALARIES where  '+st1+' and'+
        ' (PSA_DATESORTIE >="'+usdatetime(daterech)+'" or PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR'+
        ' PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+ usdatetime(daterech)+'")'
      else
        sqlsal := 'select PSA_SALARIE, PSA_PRENOM, PSA_LIBELLE from SALARIES where'+
        ' (PSA_DATESORTIE >="'+usdatetime(daterech)+'" or PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR'+
        ' PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+usdatetime(daterech)+'")';
      // constitution de la tob salarié
      Q4:= Opensql(sqlsal, true);
      while not Q4.EOF do
      begin  // d4
        T1 := Tob.Create('Salarié',Tob_saltypepop,-1);
        T1.AddChampSup('SALARIE',False);
        T1.PutValue('SALARIE',Q4.findfield('PSA_SALARIE').asstring);
        T1.AddChampSup('NOM',False);
        T1.PutValue('NOM',Q4.findfield('PSA_LIBELLE').asstring);
        T1.AddChampSup('PRENOM',False);
        T1.PutValue('PRENOM',Q4.findfield('PSA_PRENOM').asstring);
        T1.AddChampSup('POPULATION',False);
        T1.PutValue('POPULATION',population);
        T1.AddChampSup('LIBELLEPOP',False);
        T1.PutValue('LIBELLEPOP',libellepop);
        Q4.next
      end;   // e4
      ferme(Q4);
      Tob_TempValIdent := Tob_ValIdent.FindNext([''],[''],False);
    end;  // e3
    QQ.next;
  end; // e1
  ferme(QQ);
  if assigned(Tob_valident) then FreeAndNil(Tob_valident);
  if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
  if assigned(Tob_paieparim) then FreeAndNil(Tob_paieparim);
  result:= tob_saltypepop;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Recherche s'il existe plusieurs populations pour un salarié et
Suite ........ : un type de population
Mots clefs ... :
*****************************************************************}
function pgdoublonpopsal(salarie, typepop : string): boolean;
var
  sql1, sql3,sql4,st1, nodossier, population, codepop, sti, numident,nomchamp, pred: String;
  QQ, Q3: TQuery; // , Q0
  tob_Identsal, Tids , tob_valident,tob_paieparim , TD : TOB;
  i, nbident, nbpop: integer;
  Tob_TempValIdent, Tob_TempParentValIdent, Tob_TempFilleValIdent : TOB;
begin
  result:= false;
  nbpop := 0;
   // Recherche des champs de PAIEPARIM utilisable comme critères de population
//  Q0 := Opensql('select PAI_IDENT,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_UTILISABLEPOP <> ""', true);
  Tob_paieparim := Tob.Create('Critères population', nil, -1);
//  Tob_paieparim.LoadDetailDB('Critères population', '', '', Q0, False);
  Tob_paieparim.LoadDetailFromSQL('select PAI_IDENT,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_UTILISABLEPOP <> ""');
//  ferme(Q0);
  // lecture des codes critères concernés par le type de population
  pred := GetPredefiniPopulation(typepop);
  nodossier := pgrendNodossier();
{  // recherche des codes pop prédéfini DOSSIER
  pred := 'DOS';
  sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION ,CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
  ' and PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and PPO_TYPEPOP LIKE "%'+ typepop +'%"';
  if not existesql(sql1) then
  begin
    // recherche des codes pop prédéfini STANDARD
    sql1 := 'select PPC_POPULATION,PPO_CODEPOP from ORDREPOPULATION,CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
    ' and PPO_PREDEFINI= "STD" and PPO_TYPEPOP LIKE "%'+ typepop +'%" ';
    pred := 'STD';
    if not existesql(sql1) then
    begin
      // recherche des codes pop prédéfini CEGID
      sql1 := 'select PPC_POPULATION, PPO_CODEPOP from ORDREPOPULATION, CODEPOPULATION where PPC_CODEPOP = PPO_CODEPOP'+
      ' and PPO_PREDEFINI= "CEG" and PPO_TYPEPOP LIKE "%'+ typepop +'%" ';
      if existesql(sql1) then pred := 'CEG' else pred := '';
    end;
  end; }
  if pred = 'DOS' then
  sql1 := 'select PPO_CODEPOP, PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP from CODEPOPULATION where PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and'+
  ' PPO_TYPEPOP LIKE "%'+ typepop +'%"'
  else
  sql1 := 'select PPO_CODEPOP, PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP from CODEPOPULATION where PPO_PREDEFINI= "'+pred+'" and PPO_TYPEPOP LIKE "%'+ typepop +'%"' ;
  QQ:= Opensql(sql1, true);
  While Not QQ.Eof Do
  begin   // d1
    codepop := QQ.findfield('PPO_CODEPOP').asstring;
    nbident := QQ.findfield('PPO_NBIDENT').asinteger;
    if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
    Tob_Identsal := Tob.Create('Nom identifiants salarie', nil, -1);
    for i := 1 to nbident do
    begin  // d2
      Sti := IntToStr(i);
      numident := QQ.findfield('PPO_IDENT' + sti).asstring;

      TD := Tob_paieparim.findfirst(['PAI_IDENT'], [numident], true);
      nomchamp := TD.getvalue('PAI_PREFIX') + '_' + TD.getvalue('PAI_SUFFIX');
      Tids := Tob.Create('Identifiant', Tob_Identsal, -1);
      Tids.AddChampSup('NOMCHAMPS', False);
      Tids.PutValue('NOMCHAMPS', Nomchamp);
      Tids.AddChampSup('ORDREIDENT',False);
      Tids.PutValue('ORDREIDENT',sti);
    end; // e2
    // Lecture des populations pour ce code population
    sql3 := 'SELECT PPC_POPULATION, PPC_LIBELLE,PPC_VALIDENT1,PPC_VALIDENT2, PPC_VALIDENT3, PPC_VALIDENT4'+
    ' FROM ORDREPOPULATION WHERE PPC_CODEPOP = "' + codepop + '" ';
    Q3 := OPENSQL(sql3, True);
    if assigned(Tob_valident) then FreeAndNil(Tob_valident);
    Tob_ValIdent := Tob.Create('Valeur identifiants', nil, -1);
    Tob_ValIdent.LoadDetailDB('Valeur identifiants', '', '', Q3, False);
    ferme(Q3);
    Tob_TempValIdent := Tob_ValIdent.FindFirst([''],[''],False);
    While Tob_TempValIdent <> nil do 
    begin   // d3
      population := Tob_TempValIdent.GetString('PPC_POPULATION');
      // constitution de la clause WHERE
      st1 := '';
      //On crée une tob parente pour la tob temporaire, car c'est la structure
      // que la fonction PGConstitutionwhere attends
      Tob_TempParentValIdent := Tob.Create('Tob mère de Valeur identifiants', nil, -1);
      Tob_TempFilleValIdent :=  Tob.Create('Tob fille de Valeur identifiants', Tob_TempParentValIdent, -1);
      Tob_TempFilleValIdent.Dupliquer(Tob_TempValIdent, False, True);
      if Assigned(tob_identsal) then
        st1 := PGConstitutionwhere(Tob_identsal, Tob_TempParentValIdent);//Tob_Valident);
      // On libère la tob
      FreeAndNil(Tob_TempParentValIdent);
      if st1 <> '' then
      sql4 := 'select PSA_SALARIE from SALARIES where PSA_SALARIE = "'+salarie+'"  and '+st1+''
      else
      sql4 := 'select PSA_SALARIE from SALARIES where PSA_SALARIE = "'+salarie+'"';
      if  existesql(sql4) then
      nbpop := nbpop + 1;
      Tob_TempValIdent := Tob_ValIdent.FindNext([''],[''],False);
    end; // e3
    QQ.next;
  end; // e1
  ferme(QQ);
  if assigned(Tob_valident) then FreeAndNil(Tob_valident);
  if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
  if assigned(Tob_Paieparim) then FreeAndNil(Tob_paieparim);
  if nbpop > 1 then result:= true;
end;

{function GetParametre(CodeParametre, Module, Champ : String; TypeParametre : String = 'POP') : String;
Var
  Qry : TQuery;
begin
  Qry:=OpenSQL('SELECT PGP_PGVALCHAMP FROM PGPARAMETRES '
              +'where ##PGP_PREDEFINI## and PGP_TYPEPARAMETRE = "POP" '
              +' and PGP_PGPARAMETRE = "'+CodeParametre+'" '
              +' and PGP_MODULECHAMP = "'+Module+'" '
              +' and PGP_PGNOMCHAMP  = "'+Champ+'" ',True,1) ;
  if Not Qry.EOF then
  BEGIN
    Result := Qry.Fields[0].AsString;
  END ;
  Ferme(Qry) ;
end;  }

{function GetParametre(Codepop, typepop, Champ : String) : String;
Var
  Qry : TQuery;
begin

PGO_PREDEFINI,PGO_NODOSSIER,PGO_CODEASSOCIE,
PGO_PGPARAMETRE,PGO_TYPEPARAMETRE,PGO_DATEVALIDITE


  Qry:=OpenSQL('SELECT PGP_PGVALCHAMP FROM PGPARAMETRES '
              +'where ##PGP_PREDEFINI## and PGP_TYPEPARAMETRE = "POP" '
              +' and PGP_PGPARAMETRE = "'+CodeParametre+'" '
              +' and PGP_MODULECHAMP = "'+Module+'" '
              +' and PGP_PGNOMCHAMP  = "'+Champ+'" ',True,1) ;
  if Not Qry.EOF then
  BEGIN
    Result := Qry.Fields[0].AsString;
  END ;
  Ferme(Qry) ;
end;  }


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 24/08/2007
Modifié le ... :   /  /
Description .. : Renvoie les types de populations (attention à libérer l'objet ensuite)
Mots clefs ... : 
*****************************************************************}
Function GetTypesPop : TStringList;
var Q :TQuery;
    SL : TStringList;
Begin
     SL := TStringList.Create;
     Q := OpenSQL('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE = "TPO"', True);
     While Not Q.EOF Do
     Begin
         SL.Add(Q.FindField('CO_CODE').AsString);
         Q.Next;
     End;
     Ferme(Q);
     Result := SL;
End;


{ TUpdateIdemPop }

constructor TUpdateIdemPop.Create;
begin
  TobParametres := TOB.Create('Tob des parametres',nil,-1);
  TobSalarie  := TOB.Create('Tob des salariés et de leurs données',nil,-1);
  ListeParametresCharges := TStringList.Create;
  ListeParametresCharges.Clear;
  ListeTypes := GetTypesPop; //PT2
end;


destructor TUpdateIdemPop.Destroy;
begin
  if assigned(TobParametres) then TobParametres.Free;
  if assigned(TobSalarie) then TobSalarie.Free;
  ListeParametresCharges.Free;
  inherited;
  If Assigned(ListeTypes) Then FreeAndNil(ListeTypes); //PT2
end;


function TUpdateIdemPop.MajSALARIEPOPUL(typepop : string; daterech : Tdatetime; ListUpdatesChamps : Array of String) : TStringList;
begin
  if isIncludeInPopulParam(ListUpdatesChamps, typepop) then
    result := MajSALARIEPOPUL(typepop, daterech)
  else
    result := TStringList.Create;
end;


function TUpdateIdemPop.MajSALARIEPOPUL(typepop : string; daterech : Tdatetime) : TStringList;
var
  ListSal, ListReal, ListRealFille, ListSalAssoc : TOB;
  GUid,  predefini : String;   //  nodossier, sql1,
  indexSal, indexTypePop : Integer;
  TempTypePop, ListeTypePop : String;
  TmpStringList : TStringList;
  {$IFDEF PRESENCE}GP : TGestionPresence;{$ENDIF}
begin
  if typepop = '<<Tous>>' then
  begin
    result := TStringList.Create;
    for indexTypePop := 0 to ListeTypes.Count-1 do //PT2
    begin
      TmpStringList := MajSALARIEPOPUL(ListeTypes[indexTypePop],daterech); //PT2
      result.AddStrings(TmpStringList);
      FreeAndNil(TmpStringList);
    end;
  end else if pos(';',typepop) > 0 then
  begin  //Plusieurs types ont été passé, on les executes 1 à 1
    result := TStringList.Create;
    ListeTypePop := typepop;
    TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
    While (TempTypePop <> '') or (ListeTypePop <> '') do
    begin
      if TempTypePop <> '' then
      begin
        TmpStringList := MajSALARIEPOPUL(TempTypePop,daterech);
        result.AddStrings(TmpStringList);
        FreeAndNil(TmpStringList);
      end;
      TempTypePop := Trim(ReadTokenPipe(ListeTypePop, ';'));
    end;
  end else begin
    //Pour le type 'HAB', on ne fait pas la mise à jour
    if typepop = 'HAB' then
    begin
      result := TStringList.Create;
      exit;
    //PT1 - Début
{$IFDEF PRESENCE}
    end
    Else If TypePop = 'PRE' Then
    Begin
          // On positionne tous les compteurs à recalculer
          CompteursARecalculer(DateRech);
{$ENDIF}          
    End;
    //PT1 - Fin
    //On fait la liste des salariés qui appartiennent à l'association
    //Pour que les salariés qui ne feront plus partie de la liste soient bien remis à 0
    ListSalAssoc := TOB.Create('Liste des salariés associés',nil,-1);
    ListSalAssoc.LoadDetailFromSQL('Select PNA_SALARIE from SALARIEPOPUL where PNA_TYPEPOP ="'+typepop+'"');
    //On cherche le prédéfini
    predefini := GetPredefiniPopulation(typepop);
    ListSal := pgsalarietypepop(typepop, predefini, daterech);
    ListReal := TOB.Create('Liste des salariés concernés',nil,-1);
    for indexSal := 0 to ListSal.FillesCount(0)-1 do
    begin
      GUid := AglGetGuid();
      ListRealFille := TOB.Create('SALARIEPOPUL',ListReal,indexSal);
      ListRealFille.PutValue ('PNA_GUID', GUid);
      ListRealFille.PutValue ('PNA_SALARIE', ListSal.Detail[indexSal].GetString('SALARIE'));
      ListRealFille.PutValue ('PNA_TYPEPOP', typepop);
      ListRealFille.PutValue ('PNA_POPULATION', ListSal.Detail[indexSal].GetString('POPULATION'));
    end;
    FreeAndNil(ListSal);
    ListReal.SetAllModifie(True);
    //On vide la table d'association pour que les salariés qui ne font plus parti de ce type de population soient bien supprimés
    ExecuteSQL('delete from SALARIEPOPUL where PNA_TYPEPOP ="'+typepop+'"');
//    ListReal.InsertOrUpdateDB(False);
    ListReal.InsertDB(nil,False);
    //On ajoute les salariés qui étaient associés et qui ne font pas partie de la nouvelle liste
    for indexSal := 0 to ListSalAssoc.FillesCount(0)-1 do
    begin
      if not Assigned(ListReal.FindFirst(['PNA_SALARIE'],[ListSalAssoc.Detail[indexSal].GetString('PNA_SALARIE')],False)) then
      begin
        ListRealFille := TOB.Create('SALARIEPOPUL',ListReal,indexSal);
        ListRealFille.PutValue ('PNA_SALARIE', ListSalAssoc.Detail[indexSal].GetString('PNA_SALARIE'));
      end;
    end;
    FreeAndNil(ListSalAssoc);
    //Mise à jour des données salarié
    result := MajDonneesSalaries(ListReal,daterech, typepop);

    {$IFDEF PRESENCE}
    //PT1 - Début
    If TypePop = 'PRE' Then
    Begin
      GP := TGestionPresence.Create(True,True,False,False,False,0,0,True);
      for indexSal := 0 to ListReal.FillesCount(0)-1 do
      begin
          If Not GP.UpdateProfil(ListReal.Detail[indexSal].GetValue('PNA_SALARIE'), DateRech) Then
               Result.Add(Format(TraduireMemoire('Salarié %s : Impossible de déterminer le profil de présence associé.'),[ListReal.Detail[indexSal].GetString('PNA_SALARIE')]));
      end;
      If Assigned(GP) Then FreeAndNil(GP);
    End;
    //PT1 - Fin
    {$ENDIF}

    FreeAndNil(ListReal);
  end;
end;


function TUpdateIdemPop.MajDonneesSalaries(TobListSalarie: TOB; daterech : Tdatetime; TypePop : String = '') : TStringList;
{$IFDEF IDEMPOP}
var
  indexAssociation, indexSalarie : integer;
  RequeteUpdate : String;
  ValeurIdemPop : String;
  TempTobParam : TOB;
  stSalarie, stParametre, stNomChamp, stPopulation : String;
{$ENDIF}  
begin
  result := TStringList.create;
{$IFDEF IDEMPOP}
  //On charge la liste des salariés
  GetSalaries(TobListSalarie);
  //On associe leurs populations
  GetAssocSalPop;
  //Puis les parametres associés
  GetAssocPopParam(daterech);
  //Pour chaque salarié, on construit la requete de mise à jour
  for indexSalarie := 0 to TobSalarie.FillesCount(0) -1 do
  begin
    stSalarie := TobSalarie.Detail[indexSalarie].GetString('PSA_SALARIE');
    RequeteUpdate := 'UPDATE SALARIES SET ';
    //On construit la requete de mise a jour
    For indexAssociation := 0 to Length(TableauAssociation)-1 do
    begin                                                 
      if (TableauAssociation[indexAssociation][0] <> '') and ((TypePop = '') or (TableauAssociation[indexAssociation][0] = TypePop)) then
      begin //On est sur le bon type
        stPopulation := TobSalarie.Detail[indexSalarie].GetString(TableauAssociation[indexAssociation][0]);
        if TobSalarie.detail[indexSalarie].GetString(TableauAssociation[indexAssociation][1]) = 'POP' then
        begin //Ce champ est de type Idem Population
          stNomChamp := RightStr(TableauAssociation[indexAssociation][2] ,Length(TableauAssociation[indexAssociation][2])-4);
          ValeurIdemPop := '';
          //On charge les paramètres
          stParametre := TobSalarie.Detail[indexSalarie].GetString(TableauAssociation[indexAssociation][0]+'_PARAM');
          if stParametre = '' then
          begin
            result.Add(TraduireMemoire('Aucun paramètre n''est')+TraduireMemoire(' associé à la population "')+ RechDom('PGPOPULATION',stPopulation,False) +TraduireMemoire('" de type "')+TableauAssociation[indexAssociation][0]+'".');
          end else begin
            GetParametres(stParametre);
            //On cherche sa valeur
            TempTobParam := TobParametres.FindFirst(['PGP_PGPARAMETRE', 'PGP_TYPEPARAMETRE', 'PGP_PGNOMCHAMP']
                                                   ,[ TobSalarie.Detail[indexSalarie].GetString(TableauAssociation[indexAssociation][0]+'_PARAM')
                                                      , 'POP'
                                                      , stNomChamp ]
                                                   ,False);
            ValeurIdemPop := '';
            if Assigned(TempTobParam) then
            begin
              ValeurIdemPop := TempTobParam.GetString('PGP_PGVALCHAMP');
              if ValeurIdemPop = '' then
                result.Add(TraduireMemoire('Salarié ')+stSalarie+' : '+TraduireMemoire('Le paramètre "')+ RechDom('PGPARAMETRE',stParametre,False)
                          +TraduireMemoire('" associé à la population "')+ RechDom('PGPOPULATION',stPopulation,False)
                          +TraduireMemoire('" est vide pour le champ "')+stNomChamp+'".');
            end else begin //On compte les erreurs, et on les ajoute dans la tobSalarie
             result.Add(TraduireMemoire('Salarié ')+stSalarie+' : '+TraduireMemoire('Le paramètre "')+ RechDom('PGPARAMETRE',stParametre,False)
                       +TraduireMemoire('" associé à la population "')+ RechDom('PGPOPULATION',stPopulation,False)
                       +TraduireMemoire('" n''est pas défini pour le champ "')+stNomChamp+'".');
            end;
            //Puis on l'affecte
            RequeteUpdate := RequeteUpdate+TableauAssociation[indexAssociation][2]+'="'+ValeurIdemPop+'", ';
          end;
        end;
      end;
    end;
    //On enleve la dernière virgule, on rajoute la clause 'where PSA_SALARIE = "'+salarie+'"'
    // et on execute la requete
    if RequeteUpdate <> 'UPDATE SALARIES SET ' then
    begin
      RequeteUpdate := LeftStr(RequeteUpdate,Length(RequeteUpdate)-2);
      RequeteUpdate := RequeteUpdate + 'where PSA_SALARIE = "'+TobSalarie.Detail[indexSalarie].GetString('PSA_SALARIE')+'"';
      ExecuteSQL(RequeteUpdate);
    end;
  end;
{$ENDIF}
end;


procedure TUpdateIdemPop.GetSalaries(TobListSalarie : TOB);
var
  indexSal : integer;
  StrListSalarie : TStringList;
  PrefixeTableName, TableName : String;
begin
  TableName := '_SALARIE';
  if TobListSalarie.FillesCount(0) > 0 then
  begin
    StrListSalarie :=  TStringList.create;
    PrefixeTableName := TableToPrefixe(TobListSalarie.Detail[0].NomTable);
    if PrefixeTableName = '' then
    begin
      TableName := 'SALARIE';
    end;
    //Debut PT4
    if not TobListSalarie.Detail[0].FieldExists(PrefixeTableName+TableName) then
    begin
      PrefixeTableName := 'PSA';
      TableName := '_SALARIE';
    end;
    //Fin PT4
    for indexSal := 0 to TobListSalarie.FillesCount(0) -1 do
      StrListSalarie.Add(TobListSalarie.Detail[indexSal].GetString(PrefixeTableName+TableName));
    GetSalaries(StrListSalarie);
    StrListSalarie.Free;
  end;
end;


procedure TUpdateIdemPop.GetSalaries(StrListSalarie: TStringList);
var
  indexSal, indexRequeteSelectSal : integer;
  RequeteSelectSal : String;
begin
  //On construit le select des données salarié :
  RequeteSelectSal := 'Select ';
  //On parcours tous les champs qui sont concernés par les Idem Pop
  for indexRequeteSelectSal := 0 to Length(TableauAssociation)-1 do
  begin
    RequeteSelectSal := RequeteSelectSal + TableauAssociation[indexRequeteSelectSal][1]+', ';
  end;
  RequeteSelectSal := RequeteSelectSal+' PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE in (';
  //On parcours la liste des salariés
  ListSalForInClause := '';
  for indexSal := 0 to StrListSalarie.Count -1 do
  begin
    ListSalForInClause := ListSalForInClause+'"'+StrListSalarie.Strings[indexSal]+'", ';
  end;
  ListSalForInClause := LeftStr(ListSalForInClause,Length(ListSalForInClause)-2);
  RequeteSelectSal := RequeteSelectSal+ListSalForInClause+')';
  //On charge les données du(des) salarié(s)
  TobSalarie.LoadDetailFromSQL(RequeteSelectSal);
end;

{$IFDEF IDEMPOP}
procedure TUpdateIdemPop.GetParametres(CodeParametre, Module,
  TypeParametre: String);
var
  Index: Integer;
begin
  if    (not ListeParametresCharges.Find(CodeParametre+Module+TypeParametre,Index))
    and (not ListeParametresCharges.Find(CodeParametre+TypeParametre,Index))
   then
  begin
    TobParametres.LoadDetailFromSQL('select * from pgparametres where ##PGP_PREDEFINI## '
                                   +' AND PGP_PGPARAMETRE = "'+CodeParametre+'" '
                                   +' AND (PGP_MODULECHAMP = "'+Module+'" or "'+Module+'" = "" )'
                                   +' and PGP_TYPEPARAMETRE = "'+TypeParametre+'"',True);
    ListeParametresCharges.Add(CodeParametre+Module+TypeParametre);
  end;
end;


procedure TUpdateIdemPop.GetAssocSalPop();
var
  indexAssoc, indexTypePop : integer;
  Requete : String;
  TobTempAssoc, TobTempSal : TOB;
begin
  if TobSalarie.FillesCount(0) > 0 then
  begin
    Requete := 'SELECT PNA_GUID, PNA_SALARIE, PNA_TYPEPOP, PNA_POPULATION FROM SALARIEPOPUL '
             + ' WHERE PNA_SALARIE IN ('+ListSalForInClause+')';
    TobTempAssoc := TOB.Create('Tob temporaire d''association salarié - population',nil,-1);
    TobTempAssoc.LoadDetailFromSQL(Requete);
    // on charge les associations dans la tob Salarié
    for indexTypePop := 0 to ListeTypes.Count-1 do //PT2
    begin
      TobSalarie.Detail[0].AddChampSup(ListeTypes[indexTypePop],True); //PT2
    end;
    for indexAssoc := 0 to TobTempAssoc.FillesCount(0) -1 do
    begin
      TobTempSal := TobSalarie.FindFirst(['PSA_SALARIE'],[TobTempAssoc.Detail[indexAssoc].GetString('PNA_SALARIE')],False);
      TobTempSal.SetString(TobTempAssoc.Detail[indexAssoc].GetString('PNA_TYPEPOP'),TobTempAssoc.Detail[indexAssoc].GetString('PNA_POPULATION'));
    end;
    FreeAndNil(TobTempAssoc);
  end;
end;


procedure TUpdateIdemPop.GetAssocPopParam(daterech : Tdatetime);
var
  indexSal, indexTypePop, Priorite, LastPriorite : integer;
  Requete, CodeAssocie, TempPredefini : String;
  TobTempAssoc, TobTemp, TobTempParametre : TOB;
begin
  TobTempParametre := nil;
  if TobSalarie.FillesCount(0) > 0 then
  begin
    Requete := 'select * from pgparametresassoc a where  ##a.PGO_PREDEFINI##  AND PGO_TYPEPARAMETRE = "POP"'
             + 'and a.PGO_DATEVALIDITE = (Select Max(b.PGO_DATEVALIDITE) from pgparametresassoc b '
                                        + 'where b.PGO_PREDEFINI = a.PGO_PREDEFINI '
                                          + 'and b.PGO_NODOSSIER = a.PGO_NODOSSIER '
                                          + 'and b.PGO_CODEASSOCIE = a.PGO_CODEASSOCIE '
                                          + 'and b.PGO_PGPARAMETRE = a.PGO_PGPARAMETRE '
                                          + 'and b.PGO_TYPEPARAMETRE = a.PGO_TYPEPARAMETRE '
                                          + 'and b.PGO_DATEVALIDITE <= "'+USDATETIME(daterech)+'" ) '
             + 'order by a.PGO_DATECREATION ';
    TobTempAssoc := TOB.Create('Tob temporaire d''association population - parametres',nil,-1);
    TobTempAssoc.LoadDetailFromSQL(Requete);
    // on charge les associations dans la tob Salarié
    for indexTypePop := 0 to ListeTypes.Count-1 do //PT2
    begin
      TobSalarie.Detail[0].AddChampSup(ListeTypes[indexTypePop]+'_PARAM',True); //PT2
    end;
    for indexSal := 0 to TobSalarie.FillesCount(0) -1 do
    begin
      for indexTypePop := 0 to ListeTypes.Count-1 do //PT2
      begin
        CodeAssocie := TobSalarie.Detail[indexSal].GetString(ListeTypes[indexTypePop]); //PT2
        TobTemp := TobTempAssoc.FindFirst(['PGO_CODEASSOCIE','PGO_TYPEPARAMETRE'],//'PGO_PREDEFINI',
                                          [CodeAssocie,'POP'],//'',
                                          False);
        LastPriorite := 0;
        While Assigned(TobTemp) do
        begin
          //Si ce n'est pas le bon type de parametre, on passe au suivant sans
          //prendre en compte celui-ci
          if leftStr(TobTemp.GetString('PGO_PGPARAMETRE'),3) = ListeTypes[indexTypePop] then   //PT2
          begin
            //Gestion du predefini avec ordre de priorité
            TempPredefini := TobTemp.GetString('PGO_PREDEFINI');
            Priorite := 0;
            if TempPredefini = 'DOS' then
              Priorite := 3
            else if TempPredefini = 'STD' then
              Priorite := 2
            else if TempPredefini = 'CEG' then
              Priorite := 1;
            if Priorite >= LastPriorite then
            //Comme les associations sont triées par dates décroissantes
            //prendre le dernier valide (même si la priorité est le même)
            // permet de prendre le plus reçent
            begin
              TobTempParametre := TobTemp;
              LastPriorite := Priorite;
            end;
          end;
          TobTemp := TobTempAssoc.FindNext(['PGO_CODEASSOCIE','PGO_TYPEPARAMETRE'],//'PGO_PREDEFINI',
                                           [TobSalarie.Detail[indexSal].GetString(ListeTypes[indexTypePop]),'POP'],//'', //PT2
                                           False);
        end;
        if Assigned(TobTempParametre) then
        begin
          TobSalarie.Detail[indexSal].SetString(ListeTypes[indexTypePop]+'_PARAM',  //PT2
                                                TobTempParametre.GetString('PGO_PGPARAMETRE'));
          TobTempParametre := nil;
        end else begin
          TobSalarie.Detail[indexSal].SetString(ListeTypes[indexTypePop]+'_PARAM','');  //PT2
        end;
      end;
    end;
    FreeAndNil(TobTempAssoc);
  end;
end;
{$ENDIF}


function TUpdateIdemPop.MajDonneesSalariePop(Population: string;
  daterech: Tdatetime): TStringList;
var
  TempTob : Tob;
begin
  TempTob := salariepop(GetCodePop(Population),Population,daterech);
  result := MajDonneesSalaries(TempTob, daterech);
  FreeAndNil(TempTob);
end;


function TUpdateIdemPop.GetCodePop(Population: String): String;
Var
  Qry : TQuery;
begin
  Result := '';
  Qry:=OpenSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION '
              +'where ##PPC_PREDEFINI## '
              +' and PPC_CODEPOP = "'+Population+'" ',True,1) ;
  if Not Qry.EOF then Result := Qry.Fields[0].AsString;
  Ferme(Qry) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 19/04/2007
Modifié le ... : 19/04/2007
Description .. : La tob en entrée doit contenir au rang 1 une tob par salarié
Suite ........ : avec le matricule salarié dans une colonne
Suite ........ : "SALARIE", et chacune de ces tobs doit elle même
Suite ........ : contenir une tob par champ avec 2 colonnes : une
Suite ........ : colonne "CHAMPS" qui contiens le nom du champ et une
Suite ........ : colonne "VALEUR" qui contiens la valeur du champ.
Mots clefs ... :
*****************************************************************}
{
//Fonction OBSOLETE
//C'est maintenant le parametrage lui même qui est validé


function TUpdateIdemPop.ValideModifSalarie(TobSalarieAndUpdates: TOB): Boolean;
var
  indexSalarie, indexChamp, indexRequete, indexTypePop, tempNbrErr, indexerreur : Integer;
  TempRequete, MsgErrNonBloquant, stListErreurs : String;
  ListRequete, ListErreurs : TStringList;
  stModalResult : String;
begin
  Result := True;  }
//{$IFNDEF EAGLSERVER}
{  MsgErrNonBloquant := '';
  ListRequete := TStringList.Create;
  //On construit la requete de mise a jour pour chaque salarié
  for indexSalarie := 0 to TobSalarieAndUpdates.FillesCount(0) -1 do
  begin
    if TobSalarieAndUpdates.Detail[indexSalarie].FillesCount(1) > 0 then
    begin
      TempRequete := 'Update SALARIES set ';
      //l'indice passé à la fonction FillesCount est l'indice absolue jusqu'ou va l'exploration
      //sachant qu'on part du niveau de la tob appelante (ici TobSalarieAndUpdates.Detail[indexSalarie] de rang 1
      // jusqu'au rang 1)
      for indexChamp := 0 to TobSalarieAndUpdates.Detail[indexSalarie].FillesCount(1) -1 do
      begin
        TempRequete := TempRequete
                     + 'PSA_'+TobSalarieAndUpdates.Detail[indexSalarie].Detail[indexChamp].GetString('CHAMPS')
                     +' = "'
                     + TobSalarieAndUpdates.Detail[indexSalarie].Detail[indexChamp].GetString('VALEUR')
                     +'", '
      end;
      //On supprime la dernière virgule
      TempRequete := LeftStr(TempRequete, Length(TempRequete)-2);
      TempRequete := TempRequete + ' WHERE PSA_SALARIE = "'+TobSalarieAndUpdates.Detail[indexSalarie].GetString('SALARIE')+'"';
      ListRequete.Add(TempRequete);
    end;
  end;
  //On commence une transaction, et on fait les mises à jour
  BEGINTRANS;
  try
    for indexRequete := 0 to ListRequete.Count -1 do
    begin
      ExecuteSQL(ListRequete.Strings[indexRequete]);
    end;
    ListRequete.Free;
    ListErreurs  := TStringList.Create;
    //On test les populations
    for indexSalarie := 0 to TobSalarieAndUpdates.FillesCount(0) -1 do
    begin
      for indexTypePop := 0 to length(TableauTypePop)-1 do
      begin
      //Test des doublons
        if pgdoublonpopsal(TobSalarieAndUpdates.Detail[indexSalarie].GetString('SALARIE'),TableauTypePop[indexTypePop]) then
        begin
          PGIError('Les valeurs modifiées seraient à l''origine de doublons.');
          result := False;
          ROLLBACK;
          ListErreurs.Free;
          Exit;
        end;
      //Test des exclus
        if pgpopsalarie(TobSalarieAndUpdates.Detail[indexSalarie].GetString('SALARIE'),TableauTypePop[indexTypePop]) = '' then
        begin
          ListErreurs.Add('le salarié '
                         + TobSalarieAndUpdates.Detail[indexSalarie].GetString('SALARIE')
                         + ' est exclu des populations de type '
                         + TableauTypePop[indexTypePop]);
        end;
      end;
    end;
    tempNbrErr := ListErreurs.Count;
    if tempNbrErr > 0 then
      MsgErrNonBloquant := 'Des salariés vont être exclus des populations.';
    MsgErrNonBloquant := MsgErrNonBloquant + ';';
      //Simulation de Maj des IdemPop  }
//{$IFDEF IDEMPOP}
{    ListErreurs.AddStrings(Self.MajDonneesSalaries(TobSalarieAndUpdates,Date()));
    if ListErreurs.Count > tempNbrErr then
      MsgErrNonBloquant := MsgErrNonBloquant + 'Des salariés vont générer des alertes au niveau des zones de type "Idem Population".';
}
//{$ENDIF}
{    MsgErrNonBloquant := MsgErrNonBloquant + ';';

  //On annule nos modfications, elles seront effectuées par l'Update de l'AGL si elles ont été acceptées.
  Finally
   ROLLBACK;
  end;
  //Des erreurs sont survenues, on affiche un message d'alerte
  if MsgErrNonBloquant <> '' then
  begin
    MsgErrNonBloquant := MsgErrNonBloquant + 'Voulez-vous quand même valider les modifications ?';
    stListErreurs := '';
    for indexerreur := 0 to ListErreurs.count -1 do
    begin
      stListErreurs := stListErreurs + ListErreurs.Strings[indexerreur]+'|';
    end;
    //Controls
    stModalResult := '';
    stModalResult := AGLLanceFiche('PAY', 'MSGBOX_LISTALERTE', '', '', TraduireMemoire('Gestion des populations')+';'+MsgErrNonBloquant+';'+stListErreurs);
    if stModalResult = 'Non' then
      result := False;
  end;
  ListErreurs.Free;  }
//{$ENDIF}
{end;}


function TUpdateIdemPop.MajSALARIEPOPULSalarie(Salarie: String; daterech : Tdatetime; ListUpdatesChamps : Array of String) : TStringList;
begin
  if isIncludeInPopulParam(ListUpdatesChamps, 'PAI') then
    result := MajSALARIEPOPULSalarie(Salarie, daterech)
  else
    result := TStringList.Create;
end;


function TUpdateIdemPop.MajSALARIEPOPULSalarie(Salarie: String;
  daterech : Tdatetime) : TStringList;
var
  indexTypePop : Integer;
  TypePop, Population : String;
  ListSal, sal : TOB;
  TmpStrListe : TStringList;
  CanUse : Array of Boolean;
begin
  result := TStringList.Create;
  SetLength(CanUse,ListeTypes.Count); //PT2
  //On test les types de population
  for indexTypePop := 0 to ListeTypes.Count-1 do //PT2
  begin
    TypePop := ListeTypes[indexTypePop];  //PT2
    CanUse[indexTypePop] := CanUsePopulation(TypePop);
  end;

  //On vide la table d'association
  ExecuteSQL('delete from SALARIEPOPUL where PNA_SALARIE ="'+Salarie+'"');
  for indexTypePop := 0 to ListeTypes.Count-1 do  //PT2
  begin
    TypePop := ListeTypes[indexTypePop];  //PT2
    if TypePop <> 'HAB' then
    begin
      //On teste si le paramétrage est valide pour ce type de population
      if (GetPredefiniPopulation(TypePop) <> '') and (CanUse[indexTypePop]) then
      begin
        Population := pgpopsalarie(Salarie, TypePop);
        if Population = '' then
        begin
          result.Add(TraduireMemoire('Le salarié ')+Salarie+TraduireMemoire(' n''a pas de population de type ')+TypePop+'.');
        end else begin
          ExecuteSQL('INSERT INTO SALARIEPOPUL (PNA_GUID, PNA_SALARIE, PNA_TYPEPOP, PNA_POPULATION)'+
                       ' VALUES ("'+AglGetGuid()+'","'+Salarie+'","'+TypePop+'","'+Population+'")');
          {$IFDEF PRESENCE}
          //PT1 - Début
          If TypePop = 'PRE' Then
          Begin
               // On positionne tous les compteurs à recalculer pour le salarié
               CompteursARecalculer(DateRech, Salarie);

               // On réaffecte le profil adéquat en fonction de la nouvelle population
               UpdateProfilPresenceFromPop (Salarie, DateRech);
          End;
          //PT1 - Fin
          {$ENDIF}
        end;
      end;
    end;
  end;
  ListSal := TOB.Create('Les salariés',nil,-1);
  sal := Tob.create('SALARIE',ListSal,-1);
  sal.AddChampSupValeur('PSA_SALARIE',Salarie);
  TmpStrListe := MajDonneesSalaries(ListSal,daterech, typepop);
  result.AddStrings(TmpStrListe);
  FreeAndNil(TmpStrListe);
  FreeAndNil(ListSal);
end;

{ TCompteurReference }

procedure TCompteurReference.AddCritere(IDENT: integer);
var
  index : integer;
begin
  for index := 0 to Length(FCriteres) -1 do
  begin
    if IDENT = FCriteres[index] then
      exit;
  end;
  SetLength(FCriteres,Length(FCriteres) + 1);
  SetLength(FValeursCriteres,Length(FCriteres) + 1);
  FCriteres[Length(FCriteres) - 1] := IDENT;
end;

procedure TCompteurReference.AddCritereValeur(IDENT : integer; valeur: String);
var
  index : Integer;
begin
  for index := 0 to Length(FCriteres) -1 do
  begin
    if IDENT = FCriteres[index] then
    begin
      SetValCritere(IDENT, valeur);
      exit;
    end;
  end;
  SetLength(FCriteres,Length(FCriteres) + 1);
  SetLength(FValeursCriteres,Length(FCriteres));
  FCriteres[Length(FCriteres) - 1] := IDENT;
  FValeursCriteres[Length(FCriteres) - 1] := valeur;
end;

constructor TCompteurReference.Create;
begin
  SetLength(FCriteres,0);
  SetLength(FValeursCriteres,0);
end;

function TCompteurReference.CritereExiste(IDENT: integer): Boolean;
var
  index : integer;
begin
  result := False;
  for index := 0 to Length(FCriteres) -1 do
  begin
    if IDENT = FCriteres[index] then
    begin
      result := True;
      exit;
    end;
  end;
end;

destructor TCompteurReference.Destroy;
begin
  SetLength(FCriteres,0);
  SetLength(FValeursCriteres,0);
  inherited;
end;

function TCompteurReference.GetNbrReferences: Integer;
begin
  result := FCompteur;
end;

function TCompteurReference.GetValCritere(IDENT: integer): String;
var
  index : integer;
begin
  result := '';
  for index := 0 to Length(FCriteres) -1 do
  begin
    if IDENT = FCriteres[index] then
    begin
      result := FValeursCriteres[index];
      exit;
    end;
  end;
end;

procedure TCompteurReference.Increase;
begin
  Inc(FCompteur);
end;

procedure TCompteurReference.SetValCritere(IDENT : integer; valeur: String);
var
  index : integer;
begin
  for index := 0 to Length(FCriteres) -1 do
  begin
    if IDENT = FCriteres[index] then
    begin
      FValeursCriteres[index] := valeur;
      exit;
    end;
  end;
end;

{ TMatriceCompteurReference }

procedure TMatriceCompteurReference.AddCompteur(
  CriteresIDENT: array of integer; valeurs: array of String);
var
  Compteur : TCompteurReference;
  index : Integer;
begin
  Compteur := TCompteurReference.Create;
  Compteur.FTypePopulation := TypePopulation;
  Compteur.FPredefini := Predefini;
  if (length(CriteresIDENT) > 0) and (length(valeurs) = length(CriteresIDENT)) then
  begin
    for index := 0 to Length(CriteresIDENT) -1 do
    begin
      Compteur.AddCritereValeur(CriteresIDENT[index], valeurs[index]);
    end;
  end;
  SetLength(ListeCompteurs, Length(ListeCompteurs) +1 );
  ListeCompteurs[Length(ListeCompteurs)-1] := Compteur;
end;

procedure TMatriceCompteurReference.AddCritere(IDENT: integer);
var
  index, indexCritereVal, TailleInitiale : Integer;
  TabletteName, RetourTablette, PrefixeTablette, SQLT : String;
  TOBTablette : Tob;
begin
  if not CritereExiste(IDENT) then
  begin
    SetLength(FCriteres,Length(FCriteres)+1);
    FCriteres[Length(FCriteres)-1] := IDENT;
    TabletteName := FindTabletteName(IDENT);
    GetTablettePrefixeRetour(TabletteName, PrefixeTablette, RetourTablette);
    SQLT := GetTabletteSql(TabletteName,'');
    TOBTablette := TOB.Create(PrefixeToTable(PrefixeTablette),nil,-1);
    TOBTablette.LoadDetailDBFromSql(PrefixeToTable(PrefixeTablette), SQLT);
    TailleInitiale := Length(ListeCompteurs);
    if TailleInitiale = 0 then
    begin
      AddCompteur([], []);
      TailleInitiale := 1;
    end;
    for index := 0 to TailleInitiale -1 do
    begin
      ListeCompteurs[index].AddCritereValeur(IDENT,TOBTablette.Detail[0].GetString(RetourTablette));
      for indexCritereVal := 1 to (TOBTablette.FillesCount(0)-1) do
      begin
        AddCompteur(ListeCompteurs[index].FCriteres, ListeCompteurs[index].FValeursCriteres);
        ListeCompteurs[Length(ListeCompteurs)-1].AddCritereValeur(IDENT,TOBTablette.Detail[indexCritereVal].GetString(RetourTablette));
      end;
    end;
    TOBTablette.Free;
  end;
end;

procedure TMatriceCompteurReference.AddCriteres(CriteresIDENT: array of Integer);
var
  index : Integer;
begin
  for index := 0 to Length(CriteresIDENT) -1 do
  begin
    AddCritere(CriteresIDENT[index]);
  end;
end;

constructor TMatriceCompteurReference.Create;
begin
  SetLength(ListeCompteurs, 0);
  SetLength(FCriteres,0);
end;

function TMatriceCompteurReference.CritereExiste(
  IDENT: integer): Boolean;
var
  index : integer;
begin
  result := False;
  for index := 0 to Length(FCriteres) -1 do
  begin
    if IDENT = FCriteres[index] then
    begin
      result := True;
      exit;
    end;
  end;
end;

destructor TMatriceCompteurReference.Destroy;
var
  indexCompteur : Integer;
begin
  for indexCompteur := 0 to Length(ListeCompteurs) -1 do
  begin
    if assigned(ListeCompteurs[indexCompteur]) then FreeAndNil(ListeCompteurs[indexCompteur]);
  end;
  SetLength(ListeCompteurs,0);
  SetLength(FCriteres,0);
  inherited;
end;

function TMatriceCompteurReference.GetCompteur(
  ListeCriteres: array of integer;
  ListeValeurs: array of String): TCompteurReference;
var
  index, indexCritere : integer;
  Conforme : Boolean;
begin
  result := nil;
  if Length(ListeCriteres) <> Length(ListeValeurs) then exit;
  If Length(FCriteres) <> Length(ListeCriteres) then exit;
  for index := 0 to Length(ListeCompteurs) -1 do
  begin
    Conforme := True;
    //On vérifie la conformité de chaque parametres
    for indexCritere := 0 to Length(ListeCriteres)-1 do
    begin
      if CompareText(ListeCompteurs[index].GetValCritere(ListeCriteres[indexCritere]),ListeValeurs[indexCritere]) <> 0 then
      begin
        Conforme := False;
        Break;
      end;
    end;
    if Conforme = True then
    begin
      result := ListeCompteurs[index];
      exit;
    end;
  end;
end;

function TMatriceCompteurReference.IsEmpty: Boolean;
begin
  if Length(ListeCompteurs) > 0 then
    result := False
  else
    result := True;
end;

function TMatriceCompteurReference.IsValide: Boolean;
var
  index : integer;
begin
  result := True;
  for index := 0 to Length(ListeCompteurs) -1 do
  begin
    if ListeCompteurs[index].GetNbrReferences <> 1 then
    begin
      result := False;
      exit;
    end;
  end;
end;

procedure TMatriceCompteurReference.UpdateCompteurs(
  ListeCriteres: array of integer; ListeValeurs: array of String);
var
  TempCompteur : TCompteurReference;
  NewListeCriteres : array of integer;
  NewListeValeur: array of String;
  indexCritere, indexCritere2, indexCritereVal, i1 : Integer;
  TempListeValeurs, TempValeur, TabletteName, PrefixeTablette, RetourTablette, SQLT : String;
  Criterefind : Boolean;
  TOBTablette : TOB;
begin
  if Length(ListeCriteres) <> Length(ListeValeurs) then exit;
  If Length(FCriteres) < Length(ListeCriteres) then exit;
  //On gere les cas ou tous les critères ne sont pas utilisés
  if Length(FCriteres) > Length(ListeCriteres) then
  begin
    //On recherche les critères manquants
    for indexCritere := 0 to Length(FCriteres)-1 do
    begin
      Criterefind := False;
      for indexCritere2 := 0 to Length(ListeCriteres)-1 do
      begin
        if ListeCriteres[indexCritere2] = FCriteres[indexCritere] then
        begin
          Criterefind := True;
          break;
        end;
      end;
      //Si le critère est manquant
      //On lance l'update pour toutes les valeurs des critères absents
      if not Criterefind then
      begin
        //On copie la liste incomplète
        SetLength(NewListeCriteres , Length(ListeCriteres)+1);
        SetLength(NewListeValeur , Length(ListeValeurs)+1);
        for i1 := 0 to Length(ListeCriteres)-1 do
        begin
          NewListeCriteres[i1] := ListeCriteres[i1];
        end;
        for i1 := 0 to Length(ListeValeurs)-1 do
        begin
          NewListeValeur[i1] := ListeValeurs[i1];
        end;
        //On ajoute le critère
        NewListeCriteres[ Length(NewListeCriteres)-1] := FCriteres[indexCritere];
        NewListeValeur[ Length(NewListeValeur)-1] := '';
        UpdateCompteurs(NewListeCriteres, NewListeValeur );
        exit;
      end;
    end;
  end;
  //On gere les cas de valeurs '' (<<TOUS>>)
  for indexCritere := 0 to Length(ListeCriteres)-1 do
  begin
    if ListeValeurs[indexCritere] = '' then
    begin
      //On récupère la liste de toutes les valeurs possibles
      TabletteName := FindTabletteName(ListeCriteres[indexCritere]);
      GetTablettePrefixeRetour(TabletteName, PrefixeTablette, RetourTablette);
      SQLT := GetTabletteSql(TabletteName,'');
      TOBTablette := TOB.Create(PrefixeToTable(PrefixeTablette),nil,-1);
      TOBTablette.LoadDetailDBFromSql(PrefixeToTable(PrefixeTablette), SQLT);
      for indexCritereVal := 0 to (TOBTablette.FillesCount(0)-1) do
      begin
        //On copie la liste
        SetLength(NewListeValeur , Length(ListeValeurs));
        for i1 := 0 to Length(ListeValeurs)-1 do
        begin
          NewListeValeur[i1] := ListeValeurs[i1];
        end;
        //On lance l'update pour chacune de ces valeurs
        NewListeValeur[indexCritere] := TOBTablette.Detail[indexCritereVal].GetString(RetourTablette);
        UpdateCompteurs(ListeCriteres, NewListeValeur );
      end;
      TOBTablette.Free;
      exit;
    end;
  end;
  //On gere les cas des valeurs multiples
  for indexCritere := 0 to Length(ListeCriteres)-1 do
  begin
    if pos(';',ListeValeurs[indexCritere]) > 0 then
    begin
      TempListeValeurs := ListeValeurs[indexCritere];
      TempValeur := Trim(ReadTokenPipe(TempListeValeurs, ';'));
      while (TempValeur <> '') or (TempListeValeurs <> '') do
      begin
        //On copie la liste
        SetLength(NewListeValeur , Length(ListeValeurs));
        for i1 := 0 to Length(ListeValeurs)-1 do
        begin
          NewListeValeur[i1] := ListeValeurs[i1];
        end;
        //On appel la fonction d'update pour chacune des valeurs
        NewListeValeur[indexCritere] := TempValeur;
        UpdateCompteurs(ListeCriteres, NewListeValeur );
        TempValeur := Trim(ReadTokenPipe(TempListeValeurs, ';'));
      end;
      exit;
    end;
  end;
  //Cas normal : Tous les critères et 1 valeur par critère
  TempCompteur := GetCompteur(ListeCriteres,ListeValeurs);
  if Assigned(TempCompteur) then
    TempCompteur.Increase;
end;

{ TParamPopulValidator }

constructor TParamPopulValidator.Create;
Var
  IndexType, IndexPredef : Integer;
begin
  SetLength(ListeMatrice,0);
  ListeTypes := GetTypesPop; //PT2
  for IndexType := 0 to ListeTypes.Count -1 do  //PT2
  begin
    for IndexPredef := 0 to Length(TableauPredefini) -1 do
    begin
      NewMatrice(ListeTypes[IndexType], TableauPredefini[IndexPredef]); //PT2
    end;
  end;
end;

destructor TParamPopulValidator.Destroy;
var
  indexMatrice : Integer;
begin
  for indexMatrice := 0 to Length(ListeMatrice) -1 do
  begin
    if assigned(ListeMatrice[indexMatrice]) then FreeAndNil(ListeMatrice[indexMatrice]);
  end;
  SetLength(ListeMatrice,0);
  inherited;
  If Assigned(ListeTypes) Then FreeAndNil(ListeTypes); //PT2
end;

function TParamPopulValidator.GetMatrice(TypePop,
  Predefini: String): TMatriceCompteurReference;
Var
  IndexType, IndexPredef : Integer;
begin
  result := nil;
  for IndexType := 0 to ListeTypes.Count -1 do  //PT2
  begin
    for IndexPredef := 0 to Length(TableauPredefini) -1 do
    begin
      if   (Length(ListeMatrice) >= IndexType*(Length(TableauPredefini))+IndexPredef+1 )
       and (Assigned(ListeMatrice[IndexType*(Length(TableauPredefini))+IndexPredef]) )
       and (CompareText(ListeMatrice[IndexType*(Length(TableauPredefini))+IndexPredef].TypePopulation,TypePop) =0)
       and (CompareText(ListeMatrice[IndexType*(Length(TableauPredefini))+IndexPredef].Predefini,Predefini) =0) then
      begin
        result := ListeMatrice[IndexType*(Length(TableauPredefini))+IndexPredef];
        exit;
      end;
    end;
  end;
end;

function TParamPopulValidator.IsValide(TypePopulation: String): Boolean;
var
//  IndexPedef : Integer;
  TempMatrice : TMatriceCompteurReference;
begin
  //Les habilitations n'ont pas besoin d'avoir un parametrage unique et exhaustif
  //on considère donc ce type de population comme toujours valide
  if TypePopulation = 'HAB' then
  begin
    result := True;
    exit;
  end;
  result := False;
{
  for IndexPedef := Length(TableauPredefini) -1 downto 0 do
  begin
    TempMatrice := GetMatrice(TypePopulation, TableauPredefini[IndexPedef]);
    if (TempMatrice <> nil) and (not TempMatrice.IsEmpty) then
    begin
      result := TempMatrice.IsValide;
      if result then exit;
    end;
  end;
}
  TempMatrice := GetMatrice(TypePopulation, GetPredefiniPopulation(TypePopulation));
  if (TempMatrice <> nil) and (not TempMatrice.IsEmpty) then
    result := TempMatrice.IsValide;
end;

procedure TParamPopulValidator.NewMatrice(TypePop, Predefini: String);
Var
  TempMatrice : TMatriceCompteurReference;
begin
  if GetMatrice(TypePop, Predefini) <> nil then exit;
  TempMatrice := TMatriceCompteurReference.Create;
  TempMatrice.TypePopulation := TypePop;
  TempMatrice.Predefini := Predefini;
  SetLength(ListeMatrice,Length(ListeMatrice)+1);
  ListeMatrice[Length(ListeMatrice)-1] := TempMatrice;
end;

procedure TParamPopulValidator.UpdateCompteurs;
var
  IndexType, IndexPredef, NbrCritere, indexCritere : Integer;
  RequeteSQL : String;
  TOBParametrage, TempTobParam : Tob;
  TempMatrice : TMatriceCompteurReference;
  ListeCriteres : Array of Integer;
  ListeValeurs : Array of String;

  TabletteName, RetourTablette, PrefixeTablette, SQLT : String;
  TOBTablette : Tob;
  indexCritereVal : Integer;
begin
  RequeteSQL := 'SELECT PPC_PREDEFINI, PPC_VALIDENT1, PPC_VALIDENT2, PPC_VALIDENT3, PPC_VALIDENT4, '
              + 'PPO_PREDEFINI, PPO_CODEPOP, PPO_TYPEPOP, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4, PPO_NBIDENT '
              + 'FROM ORDREPOPULATION left outer join codepopulation on PPC_CODEPOP = PPO_CODEPOP '
              + 'WHERE ##PPC_PREDEFINI## AND ##PPO_PREDEFINI## ';
  TOBParametrage := TOB.Create('Tob contenant le parametrage',nil,-1);
  TOBParametrage.LoadDetailDBFromSql('Tob contenant le parametrage', RequeteSQL);
  for IndexType := 0 to ListeTypes.Count -1 do //PT2
  begin
    for IndexPredef := 0 to Length(TableauPredefini) -1 do
    begin
    //on regarde le parametrage et on rempli les compteurs correspondants
      TempTobParam := TOBParametrage.FindFirst(['PPC_PREDEFINI'],[TableauPredefini[IndexPredef]],False);
      while Assigned(TempTobParam) do
      begin
        if pos(ListeTypes[IndexType],TempTobParam.getString('PPO_TYPEPOP')) > 0 then //PT2
        begin
          NbrCritere := TempTobParam.GetInteger('PPO_NBIDENT');
          SetLength(ListeCriteres,NbrCritere);
          SetLength(ListeValeurs,NbrCritere);
          for indexCritere := 0 to NbrCritere-1 do
          begin
            ListeCriteres[indexCritere] := TempTobParam.GetInteger('PPO_IDENT'+IntToStr(indexCritere+1));
            ListeValeurs[indexCritere]  := TempTobParam.GetString('PPC_VALIDENT'+IntToStr(indexCritere+1));
            //Gestion de la valeur "<<Tous>>" dans le parametrage des populations
            //( pour les valeurs concerné des critères)
            if (ListeCriteres[indexCritere] <> 0) and (ListeValeurs[indexCritere] = '') then
            begin
              //On récupère la liste des valeurs possibles
              TabletteName := FindTabletteName(ListeCriteres[indexCritere]);
              GetTablettePrefixeRetour(TabletteName, PrefixeTablette, RetourTablette);
              SQLT := GetTabletteSql(TabletteName,'');
              TOBTablette := TOB.Create(PrefixeToTable(PrefixeTablette),nil,-1);
              TOBTablette.LoadDetailDBFromSql(PrefixeToTable(PrefixeTablette), SQLT);
              for indexCritereVal := 0 to (TOBTablette.FillesCount(0)-1) do
              begin
                ListeValeurs[indexCritere] := ListeValeurs[indexCritere] + TOBTablette.Detail[indexCritereVal].GetString(RetourTablette) +';';
              end;
              FreeAndNil(TOBTablette);
            end
          end;
          TempMatrice := GetMatrice(ListeTypes[IndexType],TableauPredefini[IndexPredef]);  //PT2
          TempMatrice.UpdateCompteurs(ListeCriteres, ListeValeurs);
        end;
        TempTobParam := TOBParametrage.FindNext(['PPC_PREDEFINI'],[TableauPredefini[IndexPredef]],False);
      end;
    end;
  end;
  TOBParametrage.Free;
end;

end.

